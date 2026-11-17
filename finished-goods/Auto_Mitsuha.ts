/**
 * ===================================================================
 * Auto_Mitsuha – Xuất kho theo tồn kho và ngưỡng xuất theo cặp theo thị
 * trường từ sheet nguồn sang sheet đích.
 * ===================================================================
 *
 * Script này tự động xuất kho theo tồn kho và ngưỡng xuất theo cặp theo
 * thị trường dựa trên cấu hình.
 * Mọi thông số đều nằm trong CONFIG → dễ tùy chỉnh, mở rộng.
 *
 */
// ==========================================
// 1. CẤU HÌNH (CONFIGURATION)
// ==========================================
const MITSUHA_CONFIG = {
  tenSheetThongSo: "2026", // Sheet chứa thông số mã hàng (kích thước/trọng lượng). Có thể đổi theo năm hoặc tên khác.
  // Tra cứu thông số theo tên khi không khớp chính xác (VD: rect/life, ce/bc/wt).
  traCuuTenMo: {
    bat: true,
    diemToiThieu: 0.55,
    tyLeTokenKhopToiThieu: 0.5,
    soTokenKhopToiDa: 3,
    toMauKichThuoc: true, // Tô nền cột E khi kích thước lấy từ tra cứu mờ theo tên
    maMauNenKichThuoc: "#DDEBF7", // Xanh nhạt — phân biệt khớp chính xác theo tên/mã
  },
  xoaDuLieuCu: true, // Xóa dữ liệu cũ trong sheet auto trước khi ghi dữ liệu mới
  locTonKho: true, // Bật chức năng lọc theo số lượng tồn kho
  tonKhoToiThieu: 5, // Ngưỡng tồn kho tối thiểu mặc định cho mọi thị trường (trừ khi ghi đè theo thị trường)
  tonKhoTheoThiTruong: {
    EU: 8, // Riêng EU yêu cầu tối thiểu 8
  },
  xuatTheoCapTheoThiTruong: {
    UE: 10, // Thị trường UE chỉ xuất khi gom đủ cặp 8 (theo nhóm Hợp đồng + SKU)
  },
  locThiTruong: ["EU"], // Lọc theo thị trường (2 chữ cái đầu của Hợp đồng). Ví dụ: ["UA", "UE", "PL"]. Để trống [] để lấy tất cả.
  locNgayNhapChiTiet: {
    ngay: [], // Ví dụ: [6, 7]. Để trống [] nghĩa là không lọc theo ngày.
    thang: [], // Ví dụ: [5]. Để trống [] nghĩa là không lọc theo tháng.
    nam: [], // Ví dụ: [2026]. Để trống [] nghĩa là không lọc theo năm.
  },
  locSKU: [], // Danh sách SKU bổ sung dạng số. Ví dụ: [912345, 923456]. Các SKU này sẽ được lấy THÊM (không thay thế bộ lọc chính).
  /** Mã/SKU không đưa vào danh sách soạn (sheet auto). Số 6–7 chữ số → loại mọi biến thể A/B; chuỗi → khớp đúng mã hàng. */
  skuKhongDuocSoan: [] as (number | string)[],
  gioiHanMaHang: 0, // Số lượng mã hàng tối đa muốn lấy để soạn (ví dụ: 100). Để 0 để lấy tất cả.
  // Lộ trình kho luôn sort theo khu A/B/C… trước; không chen cả block HĐ lên trước theo vị trí xa hơn.
  uuTienHopDongHoanTatTrongWave: true,
  // false (mặc định): cùng ô vị trí (VD A1) chỉ xếp theo SKU 6–7 số — quét sạch line trước khi sang ô kế.
  // true: trong cùng ô vị trí, ưu tiên dòng HĐ “hoàn tất trong wave” rồi mới tới SKU.
  uuTienHopDongHoanTatTrongWaveTrongCungViTri: false,
  // Cùng ô vị trí: gom liền kề theo SKU 6–7 số (A/B cùng ô vẫn gom được).
  gomHangCungSKU6LienKe: true,
  // Nền vàng cuối danh sách = dòng chưa đạt chuẩn xuất (thiếu cặp / cảnh báo), tách khỏi phần đủ điều kiện.
  maNenCanhBaoKhongDatChuan: "#FFF2CC",
  uuTienTheoMaPhanLoai: {
    bat: false, // Bật/tắt ưu tiên soạn hàng theo mã phân loại (ký tự trích từ mã hợp đồng, VD: EUS… → S)
    // Thứ tự ưu tiên mặc định. Mã đứng trước → soạn trước. Mã không có trong danh sách → nhóm "KHAC" cuối.
    thuTuMacDinh: ["S", "T", "M", "C", "F"],
    // Ghi đè thứ tự riêng cho từng thị trường (2 chữ cái đầu mã HĐ). VD: { EU: ["F", "S", "T", "M", "C"] }
    thuTuTheoThiTruong: {} as { [thiTruong: string]: string[] },
  },
  toMauSKUTrung: true, // Tô ô cột C sheet auto khi cùng SKU 6–7 số xuất hiện ≥2 lần trong danh sách auto (kể cả dòng nền vàng).
  maMauSKUTrung: "#D9E1F2", // Màu SKU trùng nội bộ trên sheet auto (xanh nhạt)
  // Sheet auto + pending — ô cột C khi trùng đúng khóa đơn pending (HĐ + mã hàng + SL đơn), không chỉ trùng SKU 6 số
  toMauSKUAutoTrungPending: true,
  maMauNenSKUTrungPending: "#FFA500",
  maMauChuSKUTrungPending: "#FFFFFF",
  maMauHoanTatDonCoPending: "#E2F0D9", // Màu tô dòng khi pending + wave hiện tại vừa đủ hoàn tất đơn hàng
  /**
   *
   *
   *
   */
  sheetKhongDatDieuKienXuat: {
    bat: false,
    tenSheet: "Mitsuha-KhongDatDieuKienXuat",
  },
  // Cùng layout A→R nhưng nghĩa cột H/K khác nhau:
  // Pre-packing-list: H = SL đơn hàng, J = đã xuất trước, K = SL cần soạn.
  // Xuất công: H = SL xuất hàng, J = đã xuất trước, K = SL thùng (H luôn = K).
  hangRot: {
    bat: true,
    tenSheetPrePacking: "Pre-packing-list",
    tenSheetXuatCong: "Xuất công",
    nhanViTri: "RỚT",
    duaLenDauDanhSach: true,
    cot: {
      hopDong: "D",
      maHang: "E",
      soLuongXuatDi: "H",
      soLuongDaXuat: "J",
      soLuongSoan: "K",
      viTri: "R",
    },
  },
  cot: {
    nguon: {
      tenSheet: "Thành phẩm", // Sheet tổng hợp tồn kho
      dongBatDau: 2,

      // --- Các cột CỐ ĐỊNH (đã biết từ quy trình Xuất Kho) ---
      maHang: "C",
      slDonHang: "F",
      hopDong: "H",
      viTri: "I",
      tongNhap: "Q",
      chuaHoanThanh: "P", // Cột chưa hoàn thành sản xuất
      invoiceNo: "Z",

      // --- Các cột XUẤT KHO (để tự tính tồn kho chính xác) ---
      cotsXuatKho: ["S", "U", "W"],

      // --- Các cột NGÀY NHẬP KHO (Để lọc theo ngày) ---
      cotsNgayNhap: ["J", "L", "N"], // Ngày nhập kho 1, 2, 3
      cotsNhapKho: ["K", "M", "O"], // Số lượng nhập kho 1, 2, 3

      // --- Các cột DỰ ĐOÁN ---
      // MỞ RỘNG: Bạn có thể thay đổi chữ cái cột dưới đây cho khớp với thực tế của sheet "成品出入库明细  汇总表"
      tenHang: "D", // Tên hàng
      kichThuoc: "", // Kích thước (Không dùng)
      dvt: "", // Đơn vị tính (Không dùng)
      soLo: "", // Số lô (hiện chưa rõ cột nào)
      trongLuong: "", // Trọng lượng 1 đơn vị
      theTich: "", // Thể tích 1 đơn vị
    },
    dich: {
      tenSheet: "auto",
      dongBatDau: 2, // Dòng 1 dùng làm tiêu đề
      cotCuoi: "X",
      tonKhoThucTe: "U",
      chuaHoanThanh: "V",
      ngayNhapMoiNhat: "X",
    },
    pending: {
      // Trừ tồn tạm từ Pre-packing-list (không copy tay sang pending-auto).
      tenSheet: "Pre-packing-list",
      suaDinhDangSheetNguon: false,
      dongBatDau: 2,
      hopDong: "D",
      maHang: "E",
      // Pre-packing-list: H = SL đơn hàng (khóa pending), K = SL cần soạn (trừ tồn tạm).
      slDonHang: "H",
      soLuong: "K",
      viTri: "R",
      cotToMauSKU: "E",
    },
  },
};

// Các hàm tiện ích (Helper Functions)
// @ts-ignore: Office Scripts — hiển thị cột T: số dương → âm, 0 giữ nguyên
function chuaHoanThanhAmSo(giaTri: number): number {
  if (giaTri <= 0) return 0;
  return -Number(giaTri.toFixed(4));
}

// @ts-ignore: Office Scripts — xóa nền/in đậm/màu chữ trước khi tô lại (ClearApplyTo không luôn đủ)
function datLaiDinhDangVungAuto(
  sheet: ExcelScript.Worksheet,
  rangeA1: string,
): void {
  let fmt = sheet.getRange(rangeA1).getFormat();
  fmt.getFill().clear();
  let font = fmt.getFont();
  font.setBold(false);
  font.setColor("#000000");
}

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
function lamSach(giaTri: string | number | boolean): string {
  return String(giaTri)
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "") // Xóa mọi loại khoảng trắng ẩn
    .trim();
}

// Số từ Excel: hỗ trợ dấu phẩy/chấm theo kiểu 1.234,56 hoặc 1,234.56
// @ts-ignore: Office Scripts — dùng chung với lamSach
function parseSoTuExcel(giaTri: string | number | boolean): number {
  if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
  let s = lamSach(giaTri);
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

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
function chuanHoaNgay(giaTri: string | number | boolean): string {
  let raw = giaTri;
  // Xử lý Excel serial date
  if (typeof raw === "number" && raw > 30000) {
    let jsDate = new Date(Math.round((raw - 25569) * 86400 * 1000));
    let dd = String(jsDate.getUTCDate()).padStart(2, "0");
    let mm = String(jsDate.getUTCMonth() + 1).padStart(2, "0");
    let yyyy = jsDate.getUTCFullYear();
    return dd + "/" + mm + "/" + yyyy;
  }
  return String(raw).trim();
}

// @ts-ignore: Office Scripts — khóa ngày yyyy-mm-dd để so sánh phiếu Pre / Xuất công
function ngayKeyTuGiaTriExcel(giaTri: string | number | boolean): string {
  let s = chuanHoaNgay(giaTri);
  let match = String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return "";
  return `${match[3]}-${String(Number(match[2])).padStart(2, "0")}-${String(
    Number(match[1]),
  ).padStart(2, "0")}`;
}

// Chuẩn hóa dữ liệu vị trí: bỏ ký tự nhiễu, sắp theo khu/số, format "A5 B44 C72".
// @ts-ignore: Office Scripts — dùng chung ở main() và test helpers
function chuanHoaViTri(giaTri: string | number | boolean): string {
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
    // Sắp theo khu rồi số (A15 B25, không để B25 A15) để cùng ô dễ nhìn khi soạn.
    cum.sort((a, b) => {
      if (a.khu !== b.khu) return a.khu.localeCompare(b.khu);
      return a.so - b.so;
    });
    let daThay = new Set<string>();
    let sapXep: string[] = [];
    for (let x of cum) {
      let ma = `${x.khu}${x.so}`;
      if (daThay.has(ma)) continue;
      daThay.add(ma);
      sapXep.push(ma);
    }
    return sapXep.join(" ");
  }
  return raw;
}

// @ts-ignore: Office Scripts — dùng cho pending-auto và test helpers
function taoKhoaPendingTheoNhom(
  hopDong: string | number | boolean,
  maHang: string | number | boolean,
  slDonHang?: string | number | boolean,
  _daXuatHang?: string | number | boolean,
): string {
  let hd = String(hopDong ?? "").trim();
  let mh = lamSach(maHang);
  let sl = Number(parseSoTuExcel(slDonHang ?? 0).toFixed(4));
  let thiTruong = hd.slice(0, 2).toUpperCase();
  // Không đưa "đã xuất" vào khóa để tránh lệch key khi cột H trên auto/pending
  // đã bị cập nhật theo lần chạy trước.
  return `${thiTruong}|${hd}|${mh}|${sl}`;
}

// @ts-ignore: Office Scripts — dùng cho pending-auto và test helpers
function taoKhoaPendingTheoViTri(
  khoaNhom: string,
  viTri: string | number | boolean,
): string {
  return `${khoaNhom}|${chuanHoaViTri(viTri)}`;
}

// @ts-ignore: Office Scripts — khóa hàng rớt theo HĐ + mã hàng
function taoKhoaHangRot(
  hopDong: string | number | boolean,
  maHang: string | number | boolean,
): string {
  return `${String(hopDong ?? "")
    .trim()
    .toUpperCase()}|${lamSach(maHang)}`;
}

// @ts-ignore: Office Scripts — số phiếu Excel (434 / 434.0 / "434") về cùng một chuỗi
function chuanHoaSoPhieuHangRot(giaTri: string | number | boolean): string {
  if (giaTri == null || giaTri === "") return "";
  if (typeof giaTri === "number" && isFinite(giaTri)) {
    return String(Math.round(giaTri));
  }
  let s = String(giaTri).trim();
  if (!s) return "";
  let n = parseSoTuExcel(s);
  if (n > 0 && Math.abs(n - Math.round(n)) < 1e-6) {
    return String(Math.round(n));
  }
  return s;
}

type MetaHangRotAR = {
  dongTieuDe: number;
  cotHopDong: number;
  cotMaHang: number;
  cotSoLuongSoan: number;
  cotSoLuongDaXuat: number;
  uuTienDaXuat: boolean;
};

// @ts-ignore: Office Scripts — A=0 … R=17
function chiSoCotExcelHangRot(chuCot: string): number {
  let s = String(chuCot || "")
    .toUpperCase()
    .replace(/[^A-Z]/g, "");
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

// @ts-ignore: Office Scripts — layout cố định A→R (D HĐ, E SKU, J đã xuất, K SL soạn)
function taoMetaHangRotAR(
  dongTieuDe: number,
  uuTienDaXuat: boolean,
  cot?: {
    hopDong?: string;
    maHang?: string;
    soLuongSoan?: string;
    soLuongDaXuat?: string;
  },
): MetaHangRotAR {
  return {
    dongTieuDe,
    cotHopDong: chiSoCotExcelHangRot(cot?.hopDong || "D"),
    cotMaHang: chiSoCotExcelHangRot(cot?.maHang || "E"),
    cotSoLuongSoan: chiSoCotExcelHangRot(cot?.soLuongSoan || "K"),
    cotSoLuongDaXuat: chiSoCotExcelHangRot(cot?.soLuongDaXuat || "J"),
    uuTienDaXuat,
  };
}

// @ts-ignore: Office Scripts — tìm hàng tiêu đề (cột E = SKU, cột D = mã hợp đồng)
function timDongTieuDeHangRotAR(
  rows: (string | number | boolean)[][],
  maxScan?: number,
): number {
  let gioiHan = Math.min(rows.length, maxScan && maxScan > 0 ? maxScan : 25);
  for (let r = 0; r < gioiHan; r++) {
    let row = rows[r] || [];
    let cotE = String(row[4] ?? "")
      .toLowerCase()
      .replace(/[\s\u00A0]+/g, "");
    let cotD = String(row[3] ?? "")
      .toLowerCase()
      .replace(/[\s\u00A0]+/g, "");
    if (cotE === "sku" && cotD.indexOf("hợpđồng") >= 0) return r;
    if (cotE === "sku" && cotD.indexOf("hopdong") >= 0) return r;
  }
  return 0;
}

// @ts-ignore: Office Scripts — Pre dùng K; Xuất công ưu tiên J nếu > 0, không thì K
function laySoLuongHangRotDong(
  row: (string | number | boolean)[],
  meta: MetaHangRotAR,
): number {
  let soan = parseSoTuExcel(row[meta.cotSoLuongSoan]);
  if (!meta.uuTienDaXuat) return soan;
  let daXuat = parseSoTuExcel(row[meta.cotSoLuongDaXuat]);
  return daXuat > 0 ? daXuat : soan;
}

// @ts-ignore: Office Scripts — cộng SL theo HĐ+SKU (cột D/E/J/K)
function congDonSoLuongTheoKhoaHangRot(
  rows: (string | number | boolean)[][],
  meta: MetaHangRotAR,
): { [khoa: string]: number } {
  let kq: { [khoa: string]: number } = {};
  for (let i = meta.dongTieuDe + 1; i < rows.length; i++) {
    let row = rows[i];
    if (!row) continue;
    let hd = String(row[meta.cotHopDong] ?? "").trim();
    let mh = lamSach(row[meta.cotMaHang]);
    if (!hd || !mh) continue;
    let sl = laySoLuongHangRotDong(row, meta);
    if (sl <= 0) continue;
    let khoa = taoKhoaHangRot(hd, mh);
    kq[khoa] = Number(((kq[khoa] || 0) + sl).toFixed(4));
  }
  return kq;
}

// @ts-ignore: Office Scripts — SL rớt = đã soạn trước − đã xuất công (không âm)
function tinhSlHangRotTheoNhom(
  soanTruoc: { [khoa: string]: number },
  daXuatCong: { [khoa: string]: number },
): { [khoa: string]: number } {
  let kq: { [khoa: string]: number } = {};
  for (let khoa in soanTruoc) {
    let rot = Number(
      ((soanTruoc[khoa] || 0) - (daXuatCong[khoa] || 0)).toFixed(4),
    );
    if (rot > 1e-6) kq[khoa] = rot;
  }
  return kq;
}

// @ts-ignore: Office Scripts — nhãn cột Hàng rớt (sau Q), không ghi đè vị trí
function ganNhanCotHangRot(
  nhan: string,
  slRotGan: number,
  slDong: number,
): string {
  let nhanSach = String(nhan || "RỚT").trim() || "RỚT";
  let gan = Number(slRotGan || 0);
  let sl = Number(slDong || 0);
  let phan = gan + 1e-6 < sl ? ` ${gan}/${sl}` : "";
  return `${nhanSach}${phan}`.trim();
}

// @ts-ignore: Office Scripts — yyyy-mm-dd → dd/mm/yyyy
function ngayHienThiTuNgayKey(ngayKey: string): string {
  let match = String(ngayKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return "";
  return `${match[3]}/${match[2]}/${match[1]}`;
}

// @ts-ignore: Office Scripts — yyyy-mm-dd → Excel serial (ghi số + format dd/mm/yyyy, tránh lệch text/số)
function ngayKeySangExcelSerial(ngayKey: string): number | "" {
  let match = String(ngayKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return "";
  let yyyy = Number(match[1]);
  let mm = Number(match[2]);
  let dd = Number(match[3]);
  let jsDate = new Date(Date.UTC(yyyy, mm - 1, dd));
  return Math.floor(jsDate.getTime() / 86400000) + 25569;
}

// @ts-ignore: Office Scripts — ngày nhập mới nhất từ các cột J/L/N (Excel serial hoặc "")
function layNgayNhapMoiNhatTuCacCot(
  cacGiaTriNgay: (string | number | boolean)[],
): string | number {
  let ngayKeyMoiNhat = "";
  for (let g of cacGiaTriNgay) {
    let key = ngayKeyTuGiaTriExcel(g);
    if (!key) continue;
    if (!ngayKeyMoiNhat || key > ngayKeyMoiNhat) {
      ngayKeyMoiNhat = key;
    }
  }
  return ngayKeySangExcelSerial(ngayKeyMoiNhat);
}

// @ts-ignore: Office Scripts — cột phiếu/ngày soạn gần nhất của mã rớt
function ganChuoiPhieuNgayHangRot(soPhieu: string, ngayKey: string): string {
  let ngay = ngayHienThiTuNgayKey(ngayKey);
  let phieu = String(soPhieu || "").trim();
  if (phieu && ngay) return `${phieu} — ${ngay}`;
  return phieu || ngay;
}

type KetQuaPhanTichPreXuat = {
  pendingTheoNhom: { [khoa: string]: number };
  pendingTheoViTri: { [khoa: string]: number };
  pendingTongBanDauTheoNhom: { [khoa: string]: number };
  pendingRowsMeta: { soDongSheet: number; khoaPending: string; sku6: string }[];
  tongPendingTrongSheet: number;
  soDongPendingHopLe: number;
  slHangRotTheoNhom: { [khoa: string]: number };
  thongTinHangRotTheoNhom: {
    [khoa: string]: { soPhieu: string; ngayKey: string };
  };
};

// @ts-ignore: Office Scripts — phiếu mới nhất chưa lên xe = pending; đã lên Xuất công vẫn trừ SL xuất (Thành phẩm chưa trừ)
function phanTichPendingVaHangRotTuBangAR(
  rowsPre: (string | number | boolean)[][],
  dongTieuDePre: number,
  rowsXuat: (string | number | boolean)[][],
  dongTieuDeXuat: number,
  cot?: {
    hopDong?: string;
    maHang?: string;
    slDonHang?: string;
    soLuongSoan?: string;
    soLuongDaXuat?: string;
    viTri?: string;
  },
): KetQuaPhanTichPreXuat {
  let cotNgay = 0;
  let cotPhieu = 1;
  let cotHd = chiSoCotExcelHangRot(cot?.hopDong || "D");
  let cotMh = chiSoCotExcelHangRot(cot?.maHang || "E");
  let cotSlDon = chiSoCotExcelHangRot("H");
  let cotSlSoan = chiSoCotExcelHangRot(cot?.soLuongSoan || "K");
  let cotSlDaXuat = chiSoCotExcelHangRot(cot?.soLuongDaXuat || "J");
  let cotVt = chiSoCotExcelHangRot(cot?.viTri || "R");

  type DongGom = {
    ngayKey: string;
    soPhieu: string;
    hd: string;
    mh: string;
    slDon: string | number | boolean;
    sl: number;
    viTri: string;
    soDongSheet: number;
  };

  let docDong = (
    rows: (string | number | boolean)[][],
    dongTieuDe: number,
    uuTienDaXuat: boolean,
  ): DongGom[] => {
    let ds: DongGom[] = [];
    for (let i = dongTieuDe + 1; i < rows.length; i++) {
      let row = rows[i];
      if (!row) continue;
      let hd = String(row[cotHd] ?? "").trim();
      let mh = lamSach(row[cotMh]);
      if (!hd || !mh) continue;
      let soan = parseSoTuExcel(row[cotSlSoan]);
      let daXuat = parseSoTuExcel(row[cotSlDaXuat]);
      let slXuatDiCotH = parseSoTuExcel(row[cotSlDon]);
      let sl: number;
      if (uuTienDaXuat) {
        // Xuất công: H = SL xuất (= K thùng). Không dùng H của Pre (H Pre = SL đơn).
        if (slXuatDiCotH > 0) sl = slXuatDiCotH;
        else if (soan > 0) sl = soan;
        else sl = daXuat;
      } else {
        // Pre-packing-list: K = SL cần soạn.
        sl = soan;
      }
      if (sl <= 0) continue;
      ds.push({
        ngayKey: ngayKeyTuGiaTriExcel(row[cotNgay]),
        soPhieu: chuanHoaSoPhieuHangRot(row[cotPhieu]),
        hd,
        mh,
        slDon: row[cotSlDon],
        sl: Number(sl.toFixed(4)),
        viTri: chuanHoaViTri(row[cotVt]),
        soDongSheet: i + 1,
      });
    }
    return ds;
  };

  let dongPre = docDong(rowsPre, dongTieuDePre, false);
  let dongXuat = docDong(rowsXuat || [], dongTieuDeXuat, true);

  let slXuatTheoPhieuHdSku: { [khoa: string]: number } = {};
  let slXuatTheoPhieuSku6: { [khoa: string]: number } = {};
  for (let d of dongXuat) {
    let k = `${d.soPhieu}|${taoKhoaHangRot(d.hd, d.mh)}`;
    slXuatTheoPhieuHdSku[k] = Number(
      ((slXuatTheoPhieuHdSku[k] || 0) + d.sl).toFixed(4),
    );
    let sku6 = chuanHoaSKU6So(d.mh) || lamSach(d.mh);
    if (d.soPhieu && sku6) {
      let k6 = `${d.soPhieu}|${sku6}`;
      slXuatTheoPhieuSku6[k6] = Number(
        ((slXuatTheoPhieuSku6[k6] || 0) + d.sl).toFixed(4),
      );
    }
  }

  let laySlXuatCuaDongPre = (d: DongGom, kChinh: string): number => {
    let theoHd = slXuatTheoPhieuHdSku[kChinh] || 0;
    if (theoHd > 1e-6) return theoHd;
    let sku6 = chuanHoaSKU6So(d.mh) || lamSach(d.mh);
    if (!d.soPhieu || !sku6) return 0;
    return slXuatTheoPhieuSku6[`${d.soPhieu}|${sku6}`] || 0;
  };

  let maxNgayPre = "";
  for (let d of dongPre) {
    if (d.ngayKey && d.ngayKey > maxNgayPre) maxNgayPre = d.ngayKey;
  }

  let gomPreTheoPhieuHdSku: { [khoa: string]: DongGom } = {};
  for (let d of dongPre) {
    let k = d.soPhieu
      ? `${d.soPhieu}|${taoKhoaHangRot(d.hd, d.mh)}`
      : `__DONG__${d.soDongSheet}`;
    let cu = gomPreTheoPhieuHdSku[k];
    if (!cu || d.sl > cu.sl) gomPreTheoPhieuHdSku[k] = d;
  }

  let pendingTheoNhom: { [khoa: string]: number } = {};
  let pendingTheoViTri: { [khoa: string]: number } = {};
  let pendingTongBanDauTheoNhom: { [khoa: string]: number } = {};
  let pendingRowsMeta: {
    soDongSheet: number;
    khoaPending: string;
    sku6: string;
  }[] = [];
  let tongPendingTrongSheet = 0;
  let soDongPendingHopLe = 0;
  let slHangRotTheoNhom: { [khoa: string]: number } = {};
  let thongTinHangRotTheoNhom: {
    [khoa: string]: { soPhieu: string; ngayKey: string };
  } = {};

  let ghiThongTinRot = (khoaHdSku: string, d: DongGom) => {
    let cu = thongTinHangRotTheoNhom[khoaHdSku];
    if (
      !cu ||
      d.ngayKey > cu.ngayKey ||
      (d.ngayKey === cu.ngayKey && d.soPhieu > cu.soPhieu)
    ) {
      thongTinHangRotTheoNhom[khoaHdSku] = {
        soPhieu: d.soPhieu,
        ngayKey: d.ngayKey,
      };
    }
  };

  let congPending = (d: DongGom, slGiu?: number) => {
    let sl = slGiu == null ? d.sl : Number(slGiu.toFixed(4));
    if (sl <= 1e-6) return;
    let khoaPending = taoKhoaPendingTheoNhom(d.hd, d.mh, d.slDon);
    pendingTongBanDauTheoNhom[khoaPending] = Number(
      ((pendingTongBanDauTheoNhom[khoaPending] || 0) + sl).toFixed(4),
    );
    pendingTheoNhom[khoaPending] = Number(
      ((pendingTheoNhom[khoaPending] || 0) + sl).toFixed(4),
    );
    pendingRowsMeta.push({
      soDongSheet: d.soDongSheet,
      khoaPending,
      sku6: chuanHoaSKU6So(d.mh),
    });
    if (d.viTri) {
      let khoaVt = taoKhoaPendingTheoViTri(khoaPending, d.viTri);
      pendingTheoViTri[khoaVt] = Number(
        ((pendingTheoViTri[khoaVt] || 0) + sl).toFixed(4),
      );
    }
    tongPendingTrongSheet = Number((tongPendingTrongSheet + sl).toFixed(4));
    soDongPendingHopLe++;
  };

  let tapDaGiuPhieuHdSku = new Set<string>();

  for (let kPhieu in gomPreTheoPhieuHdSku) {
    let d = gomPreTheoPhieuHdSku[kPhieu];
    let khoaHdSku = taoKhoaHangRot(d.hd, d.mh);
    let slXuat = laySlXuatCuaDongPre(d, kPhieu);
    let daXuatSkuNay = slXuat > 1e-6;
    let laPhieuMoiNhat = !maxNgayPre || d.ngayKey === maxNgayPre;

    if (daXuatSkuNay) {
      if (slXuat > 1e-6) {
        congPending(d, slXuat);
        tapDaGiuPhieuHdSku.add(kPhieu);
      }
      let rot = Number((d.sl - slXuat).toFixed(4));
      if (rot > 1e-6) {
        slHangRotTheoNhom[khoaHdSku] = Number(
          ((slHangRotTheoNhom[khoaHdSku] || 0) + rot).toFixed(4),
        );
        ghiThongTinRot(khoaHdSku, d);
      }
      continue;
    }

    if (laPhieuMoiNhat) {
      congPending(d);
      continue;
    }

    let cu = slHangRotTheoNhom[khoaHdSku] || 0;
    if (d.sl > cu) slHangRotTheoNhom[khoaHdSku] = d.sl;
    ghiThongTinRot(khoaHdSku, d);
  }

  let maxNgayXuat = "";
  for (let d of dongXuat) {
    if (d.ngayKey && d.ngayKey > maxNgayXuat) maxNgayXuat = d.ngayKey;
  }
  let gomXuatTheoPhieuHdSku: { [khoa: string]: DongGom } = {};
  for (let d of dongXuat) {
    let k = d.soPhieu
      ? `${d.soPhieu}|${taoKhoaHangRot(d.hd, d.mh)}`
      : `__XUAT__${d.soDongSheet}`;
    let cu = gomXuatTheoPhieuHdSku[k];
    if (!cu) {
      gomXuatTheoPhieuHdSku[k] = { ...d };
    } else {
      cu.sl = Number((cu.sl + d.sl).toFixed(4));
      if (d.ngayKey > cu.ngayKey) cu.ngayKey = d.ngayKey;
    }
  }
  for (let kXuat in gomXuatTheoPhieuHdSku) {
    if (tapDaGiuPhieuHdSku.has(kXuat)) continue;
    let d = gomXuatTheoPhieuHdSku[kXuat];
    if (maxNgayXuat && d.ngayKey && d.ngayKey !== maxNgayXuat) continue;
    congPending(d);
  }

  return {
    pendingTheoNhom,
    pendingTheoViTri,
    pendingTongBanDauTheoNhom,
    pendingRowsMeta,
    tongPendingTrongSheet,
    soDongPendingHopLe,
    slHangRotTheoNhom,
    thongTinHangRotTheoNhom,
  };
}

// @ts-ignore: Office Scripts — đọc pending từ bảng A→R (Pre-packing-list)
function napPendingTuBangAR(
  rows: (string | number | boolean)[][],
  dongTieuDe: number,
  cot?: {
    hopDong?: string;
    maHang?: string;
    slDonHang?: string;
    soLuong?: string;
    viTri?: string;
  },
): {
  pendingTheoNhom: { [khoa: string]: number };
  pendingTheoViTri: { [khoa: string]: number };
  pendingTongBanDauTheoNhom: { [khoa: string]: number };
  pendingRowsMeta: { soDongSheet: number; khoaPending: string; sku6: string }[];
  tongPendingTrongSheet: number;
  soDongPendingHopLe: number;
} {
  let ketQua = phanTichPendingVaHangRotTuBangAR(rows, dongTieuDe, [], 0, {
    hopDong: cot?.hopDong,
    maHang: cot?.maHang,
    soLuongSoan: cot?.soLuong,
    viTri: cot?.viTri,
  });
  return {
    pendingTheoNhom: ketQua.pendingTheoNhom,
    pendingTheoViTri: ketQua.pendingTheoViTri,
    pendingTongBanDauTheoNhom: ketQua.pendingTongBanDauTheoNhom,
    pendingRowsMeta: ketQua.pendingRowsMeta,
    tongPendingTrongSheet: ketQua.tongPendingTrongSheet,
    soDongPendingHopLe: ketQua.soDongPendingHopLe,
  };
}

// @ts-ignore: Office Scripts — helper thuần để trừ tồn/đơn hàng tạm theo pending-auto
function rutSoLuongPending(
  tonKhoTruocPending: number,
  khoaNhom: string,
  viTri: string | number | boolean,
  pendingTheoNhom: { [khoa: string]: number },
  pendingTheoViTri: { [khoa: string]: number },
): { tonKhoSauPending: number; soLuongPendingDaTru: number } {
  let tonConLai = Number(Number(tonKhoTruocPending || 0).toFixed(4));
  if (tonConLai <= 0) {
    return { tonKhoSauPending: 0, soLuongPendingDaTru: 0 };
  }

  let pendingDaTru = 0;
  let pendingTheoNhomConLai = Number(
    Number(pendingTheoNhom[khoaNhom] || 0).toFixed(4),
  );
  if (pendingTheoNhomConLai <= 0) {
    return { tonKhoSauPending: tonConLai, soLuongPendingDaTru: 0 };
  }

  let khoaViTri = taoKhoaPendingTheoViTri(khoaNhom, viTri);
  let pendingTaiViTriRaw = Number(
    Number(pendingTheoViTri[khoaViTri] || 0).toFixed(4),
  );
  let pendingTaiViTri = Math.min(pendingTaiViTriRaw, pendingTheoNhomConLai);
  if (khoaViTri.endsWith("|")) {
    pendingTaiViTri = 0;
  }

  if (pendingTaiViTri > 0) {
    let truTaiViTri = Math.min(tonConLai, pendingTaiViTri);
    tonConLai = Number((tonConLai - truTaiViTri).toFixed(4));
    pendingDaTru = Number((pendingDaTru + truTaiViTri).toFixed(4));
    pendingTheoViTri[khoaViTri] = Number(
      (pendingTaiViTriRaw - truTaiViTri).toFixed(4),
    );
    if (pendingTheoViTri[khoaViTri] < 0) pendingTheoViTri[khoaViTri] = 0;
    pendingTheoNhomConLai = Number(
      (pendingTheoNhomConLai - truTaiViTri).toFixed(4),
    );
    if (pendingTheoNhomConLai < 0) pendingTheoNhomConLai = 0;
  }

  if (tonConLai > 0 && pendingTheoNhomConLai > 0) {
    let truTheoNhom = Math.min(tonConLai, pendingTheoNhomConLai);
    tonConLai = Number((tonConLai - truTheoNhom).toFixed(4));
    pendingDaTru = Number((pendingDaTru + truTheoNhom).toFixed(4));
    pendingTheoNhom[khoaNhom] = Number(
      (pendingTheoNhomConLai - truTheoNhom).toFixed(4),
    );
    if (pendingTheoNhom[khoaNhom] < 0) pendingTheoNhom[khoaNhom] = 0;
  } else {
    pendingTheoNhom[khoaNhom] = pendingTheoNhomConLai;
  }

  return {
    tonKhoSauPending: tonConLai < 0 ? 0 : Number(tonConLai.toFixed(4)),
    soLuongPendingDaTru: pendingDaTru < 0 ? 0 : Number(pendingDaTru.toFixed(4)),
  };
}

// @ts-ignore: Office Scripts — tránh cộng/trừ pending hai lần khi Thành phẩm đã ghi nhận xuất
function dieuChinhTonVaDaXuatSauPending(
  tongXuat: number,
  tonKhoSauPending: number,
  soLuongPendingDaTru: number,
  pendingTongNhom?: number,
): { daXuatTamTinh: number; tonKhoKhaDung: number } {
  let tx = Number(Number(tongXuat || 0).toFixed(4));
  let pendingDaTru = Number(Number(soLuongPendingDaTru || 0).toFixed(4));
  let pendingDon = Number(
    Number(
      pendingTongNhom != null && !isNaN(Number(pendingTongNhom))
        ? pendingTongNhom
        : pendingDaTru,
    ).toFixed(4),
  );
  // Tổng đã chiếm đơn = max(Thành phẩm, pending nhóm) — EU-1 đã soạn vẫn trừ khi chạy EU-2.
  let daXuatTamTinh = Number(Math.max(tx, pendingDon).toFixed(4));
  // Phần pending trùng tongXuat đã trừ tồn ở rutSoLuongPending → hoàn lại (theo dòng/vị trí).
  let pendingDaPhanAnhTrenTP = Math.min(pendingDaTru, tx);
  let tonKhoKhaDung = Number(
    (Number(tonKhoSauPending || 0) + pendingDaPhanAnhTrenTP).toFixed(4),
  );
  if (tonKhoKhaDung < 0) tonKhoKhaDung = 0;
  return { daXuatTamTinh, tonKhoKhaDung };
}

// @ts-ignore: Office Scripts — helper thuần để nhận diện nhóm HĐ+SKU có pending và được hoàn tất bởi wave hiện tại
function tinhTapNhomHoanTatCoPending(
  items: {
    khoaNhom: string;
    conLaiTruocXuat: number;
    soLuongXuat: number;
  }[],
  pendingTongBanDauTheoNhom: { [khoa: string]: number },
): Set<string> {
  let epsilon = 1e-3;
  let agg = new Map<string, { conLai: number; xuat: number }>();

  for (let it of items) {
    let cur = agg.get(it.khoaNhom);
    if (!cur) {
      cur = { conLai: 0, xuat: 0 };
      agg.set(it.khoaNhom, cur);
    }
    cur.conLai += Number(Number(it.conLaiTruocXuat).toFixed(4));
    cur.xuat += Number(Number(it.soLuongXuat).toFixed(4));
  }

  let ketQua = new Set<string>();
  for (let [khoaNhom, aggValue] of Array.from(agg.entries())) {
    let tongPending = Number(
      Number(pendingTongBanDauTheoNhom[khoaNhom] || 0).toFixed(4),
    );
    if (tongPending <= 0) continue;
    let du = aggValue.conLai - aggValue.xuat;
    if (Math.abs(du) <= epsilon) {
      ketQua.add(khoaNhom);
    }
  }

  return ketQua;
}

// @ts-ignore: Office Scripts — chuẩn hóa tên hàng để tra cứu thông số
function chuanHoaTenSanPham(giaTri: string | number | boolean): string {
  return String(giaTri ?? "")
    .toLowerCase()
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

type ThongSoTenMo = {
  kichThuoc?: string;
  trongLuong?: number;
  theTich?: number;
};

type ThongSoTheoTenEntry<T extends ThongSoTenMo> = {
  tenChuan: string;
  tokens: string[];
  thongSo: T;
  coDuLieu: boolean;
};

type TraCuuTenMoConfig = {
  diemToiThieu?: number;
  tyLeTokenKhopToiThieu?: number;
  soTokenKhopToiDa?: number;
  toMauKichThuoc?: boolean;
  maMauNenKichThuoc?: string;
};

// @ts-ignore: Office Scripts — tách token tên hàng sau chuanHoaTenSanPham
function tachTokenTenSanPham(tenChuan: string): string[] {
  if (!tenChuan) return [];
  return tenChuan.split(/\s+/).filter((t) => t.length > 0);
}

// @ts-ignore: Office Scripts — Jaccard giữa hai tập token
function tinhDiemKhopTenMo(
  queryTokens: string[],
  candidateTokens: string[],
): number {
  if (queryTokens.length === 0 || candidateTokens.length === 0) return 0;
  let setC = new Set(candidateTokens);
  let giao = 0;
  let daDemTokenQ: { [token: string]: boolean } = {};
  for (let i = 0; i < queryTokens.length; i++) {
    let token = queryTokens[i];
    if (daDemTokenQ[token]) continue;
    daDemTokenQ[token] = true;
    if (setC.has(token)) giao++;
  }
  let daDemHop: { [token: string]: boolean } = {};
  let hop = 0;
  for (let i = 0; i < queryTokens.length; i++) {
    let token = queryTokens[i];
    if (daDemHop[token]) continue;
    daDemHop[token] = true;
    hop++;
  }
  for (let i = 0; i < candidateTokens.length; i++) {
    let token = candidateTokens[i];
    if (daDemHop[token]) continue;
    daDemHop[token] = true;
    hop++;
  }
  return hop > 0 ? giao / hop : 0;
}

function demTokenTrung(
  queryTokens: string[],
  candidateTokens: string[],
): number {
  let setC = new Set(candidateTokens);
  let giao = 0;
  for (let token of queryTokens) {
    if (setC.has(token)) giao++;
  }
  return giao;
}

function layNguongTokenKhop(
  soTokenTraCuu: number,
  config: TraCuuTenMoConfig,
): number {
  let tyLe = config.tyLeTokenKhopToiThieu ?? 0.5;
  let soToiDa = config.soTokenKhopToiDa ?? 3;
  return Math.max(soToiDa, Math.ceil(soTokenTraCuu * tyLe));
}

function thongSoTenCoDuLieu(thongSo: ThongSoTenMo): boolean {
  return !!(thongSo.kichThuoc || thongSo.trongLuong || thongSo.theTich);
}

// @ts-ignore: Office Scripts — khớp mờ theo token khi không có tên trùng 100%
function timThongSoTheoTenMo<T extends ThongSoTenMo>(
  tenChuan: string,
  danhSach: ThongSoTheoTenEntry<T>[],
  config: TraCuuTenMoConfig = {},
): T | undefined {
  let queryTokens = tachTokenTenSanPham(tenChuan);
  if (queryTokens.length === 0 || danhSach.length === 0) return undefined;

  let diemToiThieu = config.diemToiThieu ?? 0.55;
  let nguongGiao = layNguongTokenKhop(queryTokens.length, config);

  let best: T | undefined;
  let bestScore = 0;
  let bestGiao = 0;
  let bestCoDuLieu = false;

  for (let item of danhSach) {
    let giao = demTokenTrung(queryTokens, item.tokens);
    if (giao < nguongGiao) continue;

    let score = tinhDiemKhopTenMo(queryTokens, item.tokens);
    if (score < diemToiThieu) continue;

    let uuTien =
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

// @ts-ignore: Office Scripts — SKU 6–7 số từ mã hàng (901019A → 901019, 9012345B → 9012345)
function chuanHoaSKU6So(giaTri: string | number | boolean): string {
  let so = String(giaTri).replace(/\D/g, "");
  return so.length >= 6 ? so : "";
}

type TapSkuKhongDuocSoan = {
  sku6: Set<string>;
  maDayDu: Set<string>;
};

// @ts-ignore — gom cấu hình skuKhongDuocSoan thành tập tra cứu
function taoTapSkuKhongDuocSoan(
  danhSach: (number | string)[],
): TapSkuKhongDuocSoan {
  let sku6 = new Set<string>();
  let maDayDu = new Set<string>();
  for (let xi = 0; xi < (danhSach || []).length; xi++) {
    let x = danhSach[xi];
    if (typeof x === "number") {
      let s = chuanHoaSKU6So(x);
      if (s) sku6.add(s);
      continue;
    }
    let mh = lamSach(String(x)).toUpperCase();
    if (!mh) continue;
    if (/^\d{6,}$/.test(mh)) {
      sku6.add(mh);
    } else {
      maDayDu.add(mh);
    }
  }
  return { sku6, maDayDu };
}

// @ts-ignore — true nếu dòng không được đưa vào soạn
function laSkuKhongDuocSoan(
  maHang: string | number | boolean,
  tap: TapSkuKhongDuocSoan,
): boolean {
  let ma = lamSach(maHang).toUpperCase();
  if (!ma) return false;
  if (tap.maDayDu.has(ma)) return true;
  let s6 = chuanHoaSKU6So(ma);
  return !!s6 && tap.sku6.has(s6);
}

function laKichThuocLayTuTraCuuTenMo(
  maHangTraCuu: string,
  tenChuan: string,
  kichThuocDangDien: string | number | boolean,
  lookupThongSo: Map<string, ThongSoTenMo>,
  lookupThongSoTheoTen: Map<string, ThongSoTenMo>,
  danhSachThongSoTheoTen: ThongSoTheoTenEntry<ThongSoTenMo>[],
  config: TraCuuTenMoConfig & { bat?: boolean },
): boolean {
  if (
    !config.bat ||
    !config.maMauNenKichThuoc ||
    config.toMauKichThuoc === false
  ) {
    return false;
  }
  if (maHangTraCuu && lookupThongSo.has(maHangTraCuu)) return false;
  if (!tenChuan || String(kichThuocDangDien ?? "").trim() === "") return false;
  if (lookupThongSoTheoTen.has(tenChuan)) return false;
  let thongSoMo = timThongSoTheoTenMo(tenChuan, danhSachThongSoTheoTen, config);
  if (!thongSoMo?.kichThuoc) return false;
  return String(kichThuocDangDien).trim() === thongSoMo.kichThuoc;
}

interface MitsuhaDongNhomNhap {
  hopDong: string | number | boolean;
  maHang: string | number | boolean;
  slDonHang: string | number | boolean;
  tongNhap: string | number | boolean;
  chuaHoanThanhBang: string | number | boolean;
}

// @ts-ignore: Office Scripts — khóa nhóm đơn theo thị trường + HĐ + SKU.
function taoKhoaNhomHopDongSKU(
  hopDong: string | number | boolean,
  maHang: string | number | boolean,
): string {
  let hd = String(hopDong ?? "").trim();
  let mh = lamSach(maHang);
  if (!hd || !mh) return "";
  let thiTruong = hd.slice(0, 2).toUpperCase();
  return `${thiTruong}|${hd}|${mh}`;
}

// @ts-ignore: Office Scripts — tính chưa hoàn thành giống main(), lấy số lớn hơn giữa bảng và tự tính.
function tinhChuaHoanThanhDong(
  slDonHang: string | number | boolean,
  tongNhap: string | number | boolean,
  chuaHoanThanhBang: string | number | boolean,
): number {
  let sldh = parseSoTuExcel(slDonHang);
  let tn = parseSoTuExcel(tongNhap);
  let chtThucTe = sldh - tn;
  if (chtThucTe < 0) chtThucTe = 0;

  let chtBang = parseSoTuExcel(chuaHoanThanhBang);
  if (chtBang < 0) chtBang = chtThucTe;

  return Number(Math.max(chtBang, chtThucTe).toFixed(4));
}

// @ts-ignore: Office Scripts — tổng chưa hoàn thành theo HĐ+SKU để chặn đơn tách nhiều đợt sản xuất.
function tongHopChuaHoanThanhTheoNhomHopDongSKU(
  danhSach: MitsuhaDongNhomNhap[],
): { [khoaNhom: string]: number } {
  let ketQua: { [khoaNhom: string]: number } = {};
  for (let dong of danhSach) {
    let khoa = taoKhoaNhomHopDongSKU(dong.hopDong, dong.maHang);
    if (!khoa) continue;
    let chuaHoanThanh = tinhChuaHoanThanhDong(
      dong.slDonHang,
      dong.tongNhap,
      dong.chuaHoanThanhBang,
    );
    ketQua[khoa] = Number(((ketQua[khoa] || 0) + chuaHoanThanh).toFixed(4));
  }
  return ketQua;
}

// @ts-ignore: Office Scripts — tên sheet Excel tối đa 31 ký tự và không chứa ký tự đặc biệt.
function taoTenSheetPrePickingList(locThiTruong: string[]): string {
  let thiTruong = (locThiTruong || [])
    .map((x) => lamSach(String(x)).toUpperCase())
    .filter((x) => x !== "")
    .join("-");
  if (!thiTruong) thiTruong = "ALL";
  let ten = `pre-picking-list-${thiTruong}`;
  ten = ten.replace(/[:\\\/\?\*\[\]]/g, "-").trim();
  if (ten.length > 31) ten = ten.slice(0, 31);
  return ten || "pre-picking-list";
}

// @ts-ignore: Office Scripts — chiều cao thiếu/0 trong thông số được quy đổi về 1.1.
function chuanHoaChieuCaoKichThuoc(giaTri: string | number | boolean): number {
  let raw = String(giaTri ?? "").trim();
  let cao = parseSoTuExcel(giaTri);
  if (raw === "" || cao === 0) return 1.1;
  return cao;
}

// @ts-ignore: Office Scripts — trích ký tự phân loại từ mã hợp đồng (EUS4001 → S, EUT1234-2 → T)
function layKyTuPhanLoaiTuHopDong(hopDong: string, thiTruong: string): string {
  let hd = String(hopDong ?? "")
    .trim()
    .toUpperCase();
  let phanConLai = hd;
  if (thiTruong !== "KHAC" && hd.startsWith(thiTruong)) {
    phanConLai = hd.slice(thiTruong.length);
  }
  let match = phanConLai.match(/[A-Z]/);
  return match ? match[0] : "KHAC";
}

// @ts-ignore: Office Scripts — chuẩn hóa danh sách ký tự phân loại (loại trùng, viết hoa)
function chuanHoaDanhSachKyTuPhanLoai(list: string[]): string[] {
  let seen: { [key: string]: boolean } = {};
  let out: string[] = [];
  for (let x of list) {
    let k = String(x ?? "")
      .trim()
      .toUpperCase();
    if (!k || seen[k]) continue;
    seen[k] = true;
    out.push(k);
  }
  return out;
}

// @ts-ignore: Office Scripts — tính mã phân loại chuẩn hóa từ mã hợp đồng
function tinhMaPhanLoaiMitsuha(
  hopDong: string,
  config: {
    thuTuMacDinh: string[];
    thuTuTheoThiTruong: { [thiTruong: string]: string[] };
  },
): string {
  let hd = String(hopDong ?? "")
    .trim()
    .toUpperCase();
  let matchTT = hd.match(/^([A-Z]{2})/);
  let thiTruong = matchTT ? matchTT[1] : "KHAC";
  let dsCauHinh = config.thuTuTheoThiTruong[thiTruong];
  let dsKyTu =
    dsCauHinh && dsCauHinh.length > 0
      ? chuanHoaDanhSachKyTuPhanLoai(dsCauHinh)
      : chuanHoaDanhSachKyTuPhanLoai(config.thuTuMacDinh);
  let raw = layKyTuPhanLoaiTuHopDong(hopDong, thiTruong);
  return dsKyTu.includes(raw) ? raw : "KHAC";
}

// @ts-ignore: Office Scripts — so sánh 2 mã phân loại theo thứ tự ưu tiên cấu hình
function soSanhMaPhanLoai(
  maPLA: string,
  maPLB: string,
  config: {
    thuTuMacDinh: string[];
    thuTuTheoThiTruong: { [thiTruong: string]: string[] };
  },
  thiTruong?: string,
): number {
  let dsCauHinh = thiTruong && config.thuTuTheoThiTruong[thiTruong];
  let dsKyTu =
    dsCauHinh && dsCauHinh.length > 0
      ? chuanHoaDanhSachKyTuPhanLoai(dsCauHinh)
      : chuanHoaDanhSachKyTuPhanLoai(config.thuTuMacDinh);
  let thuTu = [...dsKyTu, "KHAC"];
  let idxA = thuTu.indexOf(maPLA);
  let idxB = thuTu.indexOf(maPLB);
  if (idxA >= 0 || idxB >= 0) {
    if (idxA < 0) return 1;
    if (idxB < 0) return -1;
    if (idxA !== idxB) return idxA - idxB;
  }
  return maPLA.localeCompare(maPLB);
}
// Expose pure helpers for Node-based regression tests without affecting Excel execution.
const mitsuhaTestables = {
  lamSach,
  parseSoTuExcel,
  chuanHoaNgay,
  ngayKeyTuGiaTriExcel,
  ngayHienThiTuNgayKey,
  ngayKeySangExcelSerial,
  layNgayNhapMoiNhatTuCacCot,
  chuanHoaViTri,
  chuanHoaSKU6So,
  taoTapSkuKhongDuocSoan,
  laSkuKhongDuocSoan,
  taoKhoaPendingTheoNhom,
  taoKhoaPendingTheoViTri,
  chuanHoaSoPhieuHangRot,
  taoKhoaHangRot,
  chiSoCotExcelHangRot,
  taoMetaHangRotAR,
  timDongTieuDeHangRotAR,
  congDonSoLuongTheoKhoaHangRot,
  tinhSlHangRotTheoNhom,
  ganNhanCotHangRot,
  ganChuoiPhieuNgayHangRot,
  phanTichPendingVaHangRotTuBangAR,
  napPendingTuBangAR,
  rutSoLuongPending,
  dieuChinhTonVaDaXuatSauPending,
  tinhTapNhomHoanTatCoPending,
  chuanHoaTenSanPham,
  tachTokenTenSanPham,
  tinhDiemKhopTenMo,
  timThongSoTheoTenMo,
  taoKhoaNhomHopDongSKU,
  tinhChuaHoanThanhDong,
  tongHopChuaHoanThanhTheoNhomHopDongSKU,
  taoTenSheetPrePickingList,
  chuanHoaChieuCaoKichThuoc,
  layKyTuPhanLoaiTuHopDong,
  chuanHoaDanhSachKyTuPhanLoai,
  tinhMaPhanLoaiMitsuha,
  soSanhMaPhanLoai,
};

(
  globalThis as typeof globalThis & {
    __AUTO_MITSUHA_TESTABLES__?: typeof mitsuhaTestables;
  }
).__AUTO_MITSUHA_TESTABLES__ = mitsuhaTestables;

interface ProductInfo {
  trongLuong: number;
  dai: number;
  rong: number;
  cao: number;
  kichThuoc: string;
  theTich: number;
}

// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
async function main(workbook: ExcelScript.Workbook) {
  const CONFIG = MITSUHA_CONFIG;
  const COT_NGUON = CONFIG.cot.nguon;
  const COT_DICH = CONFIG.cot.dich;
  const COT_CUOI_AUTO = COT_DICH.cotCuoi || "T";
  const COT_PENDING = CONFIG.cot.pending;

  // Bộ lọc ngày nâng cao: có thể lọc riêng theo ngày/tháng/năm.
  // Quy tắc: mảng nào để trống thì bỏ qua tiêu chí đó.
  let boLocNgayChiTiet = CONFIG.locNgayNhapChiTiet;
  let dangBatLocNgayChiTiet =
    boLocNgayChiTiet.ngay.length > 0 ||
    boLocNgayChiTiet.thang.length > 0 ||
    boLocNgayChiTiet.nam.length > 0;
  let tapSkuKhongSoan = taoTapSkuKhongDuocSoan(CONFIG.skuKhongDuocSoan || []);
  let dangBatSkuKhongSoan =
    tapSkuKhongSoan.sku6.size > 0 || tapSkuKhongSoan.maDayDu.size > 0;
  let soDongBoQuaSkuKhongSoan = 0;
  let danhSachSKU = (CONFIG.locSKU || [])
    .map((x) => chuanHoaSKU6So(Number(x)))
    .filter((x) => x !== "");
  let dangBatLocSKU = danhSachSKU.length > 0;
  let dangBatLocThiTruong =
    CONFIG.locThiTruong && CONFIG.locThiTruong.length > 0;
  let dangBatUuTienPhanLoai = CONFIG.uuTienTheoMaPhanLoai?.bat === true;

  if (dangBatLocSKU && !dangBatLocThiTruong) {
    console.log(
      `LỖI CẤU HÌNH: Khi bật lọc SKU, bạn phải cấu hình locThiTruong để tránh lấy nhầm thị trường khác.`,
    );
    return;
  }

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

  // Xác định ngưỡng tồn kho tối thiểu theo thị trường (2 ký tự đầu mã hợp đồng).
  // Nếu không có cấu hình riêng thì dùng ngưỡng mặc định tonKhoToiThieu.
  let layNguongTonKhoToiThieu = (hopDong: string): number => {
    let thiTruong = hopDong.trim().slice(0, 2).toUpperCase();
    let nguongRieng =
      CONFIG.tonKhoTheoThiTruong[
        thiTruong as keyof typeof CONFIG.tonKhoTheoThiTruong
      ];
    return typeof nguongRieng === "number"
      ? nguongRieng
      : CONFIG.tonKhoToiThieu;
  };

  // Ngưỡng "xuất theo cặp" theo thị trường.
  // Ví dụ EU=8: cùng Hợp đồng + SKU phải gom đủ 8 mới cho xuất.
  let layNguongXuatTheoCap = (hopDong: string): number => {
    let thiTruong = hopDong.trim().slice(0, 2).toUpperCase();
    let nguong =
      CONFIG.xuatTheoCapTheoThiTruong[
        thiTruong as keyof typeof CONFIG.xuatTheoCapTheoThiTruong
      ];
    return typeof nguong === "number" && nguong > 0 ? nguong : 0;
  };

  // Nhận diện SKU dạng 6–7 số + hậu tố A/B (ví dụ: 901019A, 901019B, 9012345A).
  // Các SKU này phải luôn được xuất theo cặp A+B trong cùng Hợp đồng.
  let tachSKUCapAB = (
    maHang: string | number | boolean,
  ): { laSKUCapAB: boolean; skuGoc: string; hauTo: "A" | "B" | "" } => {
    let sku = lamSach(maHang).toUpperCase();
    let match = sku.match(/^(\d{6,})([AB])$/);
    if (!match) {
      return { laSKUCapAB: false, skuGoc: "", hauTo: "" };
    }
    return {
      laSKUCapAB: true,
      skuGoc: match[1],
      hauTo: match[2] as "A" | "B",
    };
  };

  // ==========================================
  // 2. ĐỌC DỮ LIỆU TỪ SHEET KHO NGUỒN
  // ==========================================
  let sourceSheet = workbook.getWorksheet(COT_NGUON.tenSheet);
  if (!sourceSheet) {
    console.log(`LỖI: Không tìm thấy sheet nguồn "${COT_NGUON.tenSheet}"`);
    return;
  }

  // Tìm dòng cuối cùng có dữ liệu của cột Mã Hàng
  let sourceLastRow =
    sourceSheet
      .getRange(COT_NGUON.maHang + "1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  if (sourceLastRow < COT_NGUON.dongBatDau) {
    console.log("Sheet kho đang trống!");
    return;
  }

  // Hàm tiện ích lấy dữ liệu cột nhanh (Hỗ trợ cột trống "")
  let getColData = (col: string): (string | number | boolean)[][] => {
    if (!col) return Array(sourceLastRow - COT_NGUON.dongBatDau + 1).fill([""]);
    return sourceSheet
      .getRange(`${col}${COT_NGUON.dongBatDau}:${col}${sourceLastRow}`)
      .getValues() as (string | number | boolean)[][];
  };

  // Đọc hàng loạt dữ liệu
  let valMaHang = getColData(COT_NGUON.maHang);
  let valHopDong = getColData(COT_NGUON.hopDong);
  let valViTri = getColData(COT_NGUON.viTri);
  let valSlDonHang = getColData(COT_NGUON.slDonHang);
  let valTongNhap = getColData(COT_NGUON.tongNhap);
  let valChuaHoanThanh = getColData(COT_NGUON.chuaHoanThanh);

  let valTenHang = getColData(COT_NGUON.tenHang);
  let valKichThuoc = getColData(COT_NGUON.kichThuoc);
  let valDVT = getColData(COT_NGUON.dvt);
  let valSoLo = getColData(COT_NGUON.soLo);
  let valTrongLuong = getColData(COT_NGUON.trongLuong);
  let valTheTich = getColData(COT_NGUON.theTich);

  // Đọc dữ liệu các cột xuất kho để tính toán tự động
  let valXuatKhos = COT_NGUON.cotsXuatKho.map((col) => getColData(col));
  // Đọc dữ liệu các cột ngày nhập kho để lọc
  let valNgayNhaps = COT_NGUON.cotsNgayNhap.map((col) => getColData(col));
  let valNhapKhos = COT_NGUON.cotsNhapKho.map((col) => getColData(col));

  let layTongNhapUuTienTuDong = (chiSoDong: number): number => {
    let tnBang = parseSoTuExcel(valTongNhap[chiSoDong][0]);
    let tongNhapTinhToan = 0;
    for (let dataNhapKho of valNhapKhos) {
      tongNhapTinhToan += parseSoTuExcel(dataNhapKho[chiSoDong][0]);
    }
    return tongNhapTinhToan !== tnBang && tongNhapTinhToan > 0
      ? Number(tongNhapTinhToan.toFixed(4))
      : Number(tnBang.toFixed(4));
  };

  let duLieuNhomNhap: MitsuhaDongNhomNhap[] = [];
  for (let i = 0; i < valMaHang.length; i++) {
    let mh = lamSach(valMaHang[i][0]);
    if (!mh) continue;
    duLieuNhomNhap.push({
      hopDong: valHopDong[i][0],
      maHang: mh,
      slDonHang: valSlDonHang[i][0],
      tongNhap: layTongNhapUuTienTuDong(i),
      chuaHoanThanhBang: valChuaHoanThanh[i][0],
    });
  }
  let chuaHoanThanhTheoNhomHopDongSKU =
    tongHopChuaHoanThanhTheoNhomHopDongSKU(duLieuNhomNhap);

  // ==========================================
  // 2.0. ĐỌC PENDING TỪ PRE-PACKING-LIST (trừ tồn tạm, không copy pending-auto)
  // ==========================================
  let laySheetKhongPhanBietHoa = (
    ten: string,
  ): ExcelScript.Worksheet | undefined => {
    let s = workbook.getWorksheet(ten);
    if (s) return s;
    let ds = workbook.getWorksheets();
    let t = String(ten || "").toLowerCase();
    for (let sh of ds) {
      if (sh.getName().toLowerCase() === t) return sh;
    }
    return undefined;
  };

  let pendingSheet = laySheetKhongPhanBietHoa(COT_PENDING.tenSheet);
  let boQuaPendingViKhongCoSheet = !pendingSheet;
  let suaDinhDangSheetPending = COT_PENDING.suaDinhDangSheetNguon === true;
  let cotToMauSKUPending = COT_PENDING.cotToMauSKU || "C";

  let pendingTheoNhom: { [khoa: string]: number } = {};
  let pendingTheoViTri: { [khoa: string]: number } = {};
  let pendingTongBanDauTheoNhom: { [khoa: string]: number } = {};
  let pendingRowsMeta: {
    soDongSheet: number;
    khoaPending: string;
    sku6: string;
  }[] = [];
  let tongPendingTrongSheet = 0;
  let soDongPendingHopLe = 0;
  let pendingLastRow = 0;
  let slHangRotTheoNhom: { [khoa: string]: number } = {};
  let thongTinHangRotTheoNhom: {
    [khoa: string]: { soPhieu: string; ngayKey: string };
  } = {};
  let soNhomHangRot = 0;
  let tongSlHangRot = 0;
  let cfgHangRot = CONFIG.hangRot;

  let rowsPrePack: (string | number | boolean)[][] = [];
  let dongTieuDePrePack = 0;
  if (pendingSheet) {
    let usedPending = pendingSheet.getUsedRange();
    if (usedPending) {
      rowsPrePack = usedPending.getValues() as (string | number | boolean)[][];
      dongTieuDePrePack = timDongTieuDeHangRotAR(rowsPrePack, 25);
      pendingLastRow = rowsPrePack.length;
    }
  }

  let rowsXuatCong: (string | number | boolean)[][] = [];
  let dongTieuDeXuatCong = 0;
  if (cfgHangRot?.bat) {
    let sheetXuat = laySheetKhongPhanBietHoa(cfgHangRot.tenSheetXuatCong);
    if (!sheetXuat) {
      console.log(
        `ℹ️ Hàng rớt: không thấy sheet "${cfgHangRot.tenSheetXuatCong}".`,
      );
    } else {
      let usedXuat = sheetXuat.getUsedRange();
      if (usedXuat) {
        rowsXuatCong = usedXuat.getValues() as (string | number | boolean)[][];
        dongTieuDeXuatCong = timDongTieuDeHangRotAR(rowsXuatCong, 25);
      }
    }
  }

  if (rowsPrePack.length > 0) {
    let phanTich = phanTichPendingVaHangRotTuBangAR(
      rowsPrePack,
      dongTieuDePrePack,
      rowsXuatCong,
      dongTieuDeXuatCong,
      {
        hopDong: COT_PENDING.hopDong,
        maHang: COT_PENDING.maHang,
        soLuongSoan: COT_PENDING.soLuong,
        viTri: COT_PENDING.viTri,
      },
    );
    pendingTheoNhom = phanTich.pendingTheoNhom;
    pendingTheoViTri = phanTich.pendingTheoViTri;
    pendingTongBanDauTheoNhom = phanTich.pendingTongBanDauTheoNhom;
    pendingRowsMeta = phanTich.pendingRowsMeta;
    tongPendingTrongSheet = phanTich.tongPendingTrongSheet;
    soDongPendingHopLe = phanTich.soDongPendingHopLe;
    slHangRotTheoNhom = phanTich.slHangRotTheoNhom;
    thongTinHangRotTheoNhom = phanTich.thongTinHangRotTheoNhom;
    for (let khoa in slHangRotTheoNhom) {
      soNhomHangRot++;
      tongSlHangRot += slHangRotTheoNhom[khoa];
    }
    tongSlHangRot = Number(tongSlHangRot.toFixed(4));
    if (cfgHangRot?.bat) {
      console.log(
        `• Hàng rớt: ${soNhomHangRot} nhóm HĐ+SKU, tổng SL ${tongSlHangRot} (phiếu cũ/không đủ trên "${cfgHangRot.tenSheetXuatCong}").`,
      );
    }
  }

  let tapKhoaPendingTuPending = new Set<string>();
  for (let pm of pendingRowsMeta) {
    if (pm.khoaPending) tapKhoaPendingTuPending.add(pm.khoaPending);
  }

  // ==========================================
  // 2.1. ĐỌC DỮ LIỆU KÍCH THƯỚC/TL TỪ SHEET THÔNG SỐ (BAYMAX) — từ dòng 2
  // ==========================================
  let lookupThongSo = new Map<string, ProductInfo>();
  let lookupThongSoTheoTen = new Map<string, ProductInfo>();
  let danhSachThongSoTheoTen: ThongSoTheoTenEntry<ProductInfo>[] = [];
  let sheetThongSo = workbook.getWorksheet(CONFIG.tenSheetThongSo);
  if (sheetThongSo) {
    let lastRowThongSo =
      sheetThongSo
        .getRange("B1048576")
        .getRangeEdge(ExcelScript.KeyboardDirection.up)
        .getRowIndex() + 1;
    if (lastRowThongSo >= 2) {
      let dataThongSo = sheetThongSo
        .getRange(`B2:H${lastRowThongSo}`)
        .getValues() as (string | number | boolean)[][];
      for (let row of dataThongSo) {
        let maHangThongSo = lamSach(row[0]);
        let tenHangThongSo = chuanHoaTenSanPham(row[1]);
        if (!maHangThongSo && !tenHangThongSo) continue;

        let trongLuongThongSo = parseSoTuExcel(row[3]);
        let dai = parseSoTuExcel(row[4]);
        let rong = parseSoTuExcel(row[5]);
        let cao = chuanHoaChieuCaoKichThuoc(row[6]);
        let theTich = (dai * rong * cao) / 1000000;

        let dStr = String(dai).replace(/\./g, ",");
        let rStr = String(rong).replace(/\./g, ",");
        let cStr = String(cao).replace(/\./g, ",");
        let kichThuocStr =
          dai === 0 && rong === 0 && cao === 0 ? "" : `${dStr}*${rStr}*${cStr}`;

        let thongSo: ProductInfo = {
          trongLuong: trongLuongThongSo,
          dai: dai,
          rong: rong,
          cao: cao,
          kichThuoc: kichThuocStr,
          theTich: theTich ? Number(theTich.toFixed(3)) : 0,
        };

        if (maHangThongSo) {
          lookupThongSo.set(maHangThongSo, thongSo);
        }
        if (tenHangThongSo) {
          let thongSoMoiCoDuLieu = thongSoTenCoDuLieu(thongSo);
          let thongSoCu = lookupThongSoTheoTen.get(tenHangThongSo);
          let thongSoCuRong = !thongSoCu || !thongSoTenCoDuLieu(thongSoCu);
          if (thongSoCuRong || thongSoMoiCoDuLieu) {
            lookupThongSoTheoTen.set(tenHangThongSo, thongSo);
            let entryMoi = {
              tenChuan: tenHangThongSo,
              tokens: tachTokenTenSanPham(tenHangThongSo),
              thongSo: thongSo,
              coDuLieu: thongSoMoiCoDuLieu,
            };
            let idx = danhSachThongSoTheoTen.findIndex(
              (item) => item.tenChuan === tenHangThongSo,
            );
            if (idx >= 0) {
              danhSachThongSoTheoTen[idx] = entryMoi;
            } else {
              danhSachThongSoTheoTen.push(entryMoi);
            }
          }
        }
      }
    }
  } else {
    console.log(
      `Lưu ý: Không tìm thấy sheet "${CONFIG.tenSheetThongSo}". Các cột Kích thước/Trọng lượng/Thể tích sẽ dùng dữ liệu mặc định.`,
    );
  }

  let traCuuThongSoTheoTen = (tenChuan: string): ProductInfo | undefined => {
    if (!tenChuan) return undefined;
    let chinhXac = lookupThongSoTheoTen.get(tenChuan);
    if (chinhXac) return chinhXac;
    if (!CONFIG.traCuuTenMo?.bat || danhSachThongSoTheoTen.length === 0) {
      return undefined;
    }
    return timThongSoTheoTenMo(
      tenChuan,
      danhSachThongSoTheoTen,
      CONFIG.traCuuTenMo,
    );
  };

  // ==========================================
  // 3. XỬ LÝ VÀ LỌC DỮ LIỆU
  // ==========================================
  let ketQuaData: (string | number | boolean)[][] = [];
  let rowsKhongDatDieuKien: (string | number | boolean)[][] = [];
  let tapSKUThieuThongSo = new Set<string>();
  let tongPendingDaApDung = 0;
  let soDongNguonDaTruPending = 0;
  let ketQuaTam: {
    dong: (string | number | boolean)[];
    khoaNhom: string;
    khoaPending: string;
    nguongTheoCap: number;
    soLuongXuat: number;
    biChanTheoCap: boolean;
    biChanCapAB: boolean;
    biChanNhomChuaDuNhap: boolean;
    biCanhBaoVang: boolean;
    khoaCapAB: string;
    hauToCapAB: "A" | "B" | "";
    hopDongKey: string;
    maHangKey: string;
    conLaiTruocXuat: number;
    lyDoKhongDat: string;
    maPhanLoai: string;
    slHangRot: number;
  }[] = [];

  for (let i = 0; i < valMaHang.length; i++) {
    let mh = lamSach(valMaHang[i][0]);
    if (!mh) continue; // Bỏ qua dòng trống không có mã hàng
    if (dangBatSkuKhongSoan && laSkuKhongDuocSoan(mh, tapSkuKhongSoan)) {
      soDongBoQuaSkuKhongSoan++;
      continue;
    }
    let skuDong = chuanHoaSKU6So(mh);
    let laSKUBoSung =
      dangBatLocSKU && !!skuDong && danhSachSKU.includes(skuDong);

    let hd = String(valHopDong[i][0]).trim();
    let vt = chuanHoaViTri(valViTri[i][0]);

    // Lọc theo Thị trường (dựa vào phần đầu của Hợp đồng - kiểm tra trong mảng)
    if (CONFIG.locThiTruong && CONFIG.locThiTruong.length > 0) {
      if (
        !CONFIG.locThiTruong.some((tt) =>
          hd.toUpperCase().startsWith(tt.toUpperCase()),
        )
      )
        continue;
    }

    // Lọc theo ngày nhập kho chi tiết
    if (dangBatLocNgayChiTiet) {
      let daTimThayNgay = false;
      for (let dataNgay of valNgayNhaps) {
        let n = chuanHoaNgay(dataNgay[i][0]);
        let khopKieuMoi = khopBoLocNgayChiTiet(n);
        if (khopKieuMoi) {
          daTimThayNgay = true;
          break;
        }
      }
      if (!daTimThayNgay && !laSKUBoSung) continue;
    }

    // Xử lý Tổng Nhập (tự động tính để dự phòng cột S bị mất công thức)
    let tn = layTongNhapUuTienTuDong(i); // Ưu tiên số tự tính nếu bảng bị lỗi

    // Tự động tính Tổng Xuất bằng cách cộng dồn các cột xuất kho
    let tongXuat = 0;
    for (let dataXuatKho of valXuatKhos) {
      let xRaw = dataXuatKho[i][0];
      let x = typeof xRaw === "number" ? xRaw : Number(lamSach(xRaw));
      if (!isNaN(x)) {
        tongXuat += x;
      }
    }

    // Tự tính tồn kho chính xác = Tổng nhập - Tổng xuất
    let tonKhoThucTe = tn - tongXuat;
    // Xử lý sai số thập phân của JS (VD: 0.1 + 0.2 = 0.30000000000000004)
    tonKhoThucTe = Number(tonKhoThucTe.toFixed(4));
    if (tonKhoThucTe < 0) tonKhoThucTe = 0;

    // Dùng hoàn toàn tồn kho tự tính (không phụ thuộc cột tồn kho trên bảng).
    let tk = tonKhoThucTe;

    // Xử lý Chưa Hoàn Thành:
    // Ưu tiên dùng cột P (thực tế vận hành), nếu không hợp lệ thì fallback sang số tự tính.
    let sldhRaw = valSlDonHang[i][0];
    let sldh = typeof sldhRaw === "number" ? sldhRaw : Number(lamSach(sldhRaw));
    if (isNaN(sldh)) sldh = 0;

    // Tự tính chưa hoàn thành: Đơn hàng - Tổng Nhập
    let chtThucTe = sldh - tn;
    if (chtThucTe < 0) chtThucTe = 0;
    let chtBangRaw = valChuaHoanThanh[i][0];
    let chtBang =
      typeof chtBangRaw === "number" ? chtBangRaw : Number(lamSach(chtBangRaw));
    if (isNaN(chtBang) || chtBang < 0) chtBang = chtThucTe;

    // Chọn số lớn hơn để không đánh giá thiếu phần còn dang dở -> tránh xuất quá tay gây kẹt kho.
    let chuaHoanThanh = Math.max(chtBang, chtThucTe);

    let daXuat = tongXuat;
    let khoaPendingNhom = taoKhoaPendingTheoNhom(hd, mh, sldh);
    let pendingTongNhom = Number(
      (pendingTongBanDauTheoNhom[khoaPendingNhom] || 0).toFixed(4),
    );
    let pendingApDung = rutSoLuongPending(
      tk,
      khoaPendingNhom,
      vt,
      pendingTheoNhom,
      pendingTheoViTri,
    );
    let soLuongPendingDaTru = pendingApDung.soLuongPendingDaTru;
    let dieuChinhPending = dieuChinhTonVaDaXuatSauPending(
      tongXuat,
      pendingApDung.tonKhoSauPending,
      soLuongPendingDaTru,
      pendingTongNhom,
    );
    tk = dieuChinhPending.tonKhoKhaDung;
    if (soLuongPendingDaTru > 0) {
      tongPendingDaApDung = Number(
        (tongPendingDaApDung + soLuongPendingDaTru).toFixed(4),
      );
      soDongNguonDaTruPending++;
    }
    let daXuatTamTinh = dieuChinhPending.daXuatTamTinh;
    let conLaiCuaDon = Number((sldh - daXuatTamTinh).toFixed(4));

    // Dữ liệu lệch nghiệp vụ → bỏ qua (không ghi dòng xuất đi = 0).
    if (tn > sldh || daXuat > sldh) {
      continue;
    }

    // Đơn đã xuất/pending hết phần còn lại → bỏ qua.
    if (conLaiCuaDon <= 0 && !laSKUBoSung) {
      continue;
    }
    if (conLaiCuaDon < 0) conLaiCuaDon = 0;

    let thiTruong = hd.trim().slice(0, 2).toUpperCase();
    let nguongTonKhoToiThieu = layNguongTonKhoToiThieu(hd);
    let nguongXuatTheoCap = layNguongXuatTheoCap(hd);
    let thongTinCapAB = tachSKUCapAB(mh);
    let khoaNhomHopDongSKU = taoKhoaNhomHopDongSKU(hd, mh);
    let nhomHopDongSKUChuaDuNhap =
      nguongXuatTheoCap > 0 &&
      (chuaHoanThanhTheoNhomHopDongSKU[khoaNhomHopDongSKU] || 0) > 0 &&
      !laSKUBoSung;

    // --- LOGIC KIỂM TRA ĐIỀU KIỆN XUẤT VÀ CHỐNG KẸT KHO ---
    // Gợi ý xuất đi tối đa bằng số lượng còn lại cần thiết của đơn hàng
    let dongXuatDi = Math.min(tk, conLaiCuaDon);

    if (CONFIG.locTonKho) {
      if (tk < nguongTonKhoToiThieu) {
        // Tồn kho dưới ngưỡng theo thị trường: CHỈ được lấy nếu đã hoàn thành xong đơn (chuaHoanThanh <= 0).
        // (Bao gồm cả ngoại lệ: Nếu số lượng đơn hàng vốn dĩ chỉ có 5, 6, 7 cái,
        // thì khi đã nhập kho đủ, chuaHoanThanh sẽ = 0, và code vẫn cho phép xuất đi bình thường).
        if (chuaHoanThanh > 0 && !laSKUBoSung) {
          continue;
        }
      }
    }

    // Gợi ý số lượng xuất đi (dongXuatDi) để tránh kẹt kho cho lần sau
    // dựa trên ngưỡng thị trường và phần chưa hoàn thành.
    if (chuaHoanThanh > 0 && chuaHoanThanh < nguongTonKhoToiThieu) {
      // Nếu số hàng đang sản xuất bị lẻ dưới ngưỡng, phải chừa lại tồn kho bù vào
      let khaoSat = tk + chuaHoanThanh - nguongTonKhoToiThieu;
      if (dongXuatDi > khaoSat) {
        dongXuatDi = khaoSat;
      }
      // Bắt buộc xuất đi tối thiểu phải >= ngưỡng cấu hình theo thị trường
      // Nếu bị kẹt kho ép số lượng xuất xuống dưới ngưỡng, thì hoàn toàn không xuất nữa, đợi lần sau
      if (dongXuatDi < nguongTonKhoToiThieu && !laSKUBoSung) {
        continue;
      }
    }

    dongXuatDi = Number(dongXuatDi.toFixed(4));
    // Không hiển thị dòng đề xuất xuất đi = 0.
    if (dongXuatDi <= 0) continue;

    // Lấy thông tin các cột khác
    let th = valTenHang[i][0];
    let kt = valKichThuoc[i][0];
    let dvt = valDVT[i][0];
    let sl = valSoLo[i][0];

    let tlRaw = valTrongLuong[i][0];
    let tl = tlRaw === "" || tlRaw == null ? 0 : tlRaw;

    let ttRaw = valTheTich[i][0];
    let tt = ttRaw === "" || ttRaw == null ? 0 : ttRaw;

    let tenHangChuan = chuanHoaTenSanPham(th);
    let thongSoTheoMa = lookupThongSo.get(mh);
    // Chỉ fallback theo tên khi không khớp mã (tránh nhầm SKU khi mã lệch ký tự ẩn đã được lamSach).
    let thongSoHang = thongSoTheoMa ?? traCuuThongSoTheoTen(tenHangChuan);
    if (!thongSoHang && mh) {
      tapSKUThieuThongSo.add(mh);
    }
    let kichThuocXuat = thongSoHang?.kichThuoc || kt || "";
    let trongLuongXuat: string | number = "";
    if (thongSoHang) {
      let v = thongSoHang.trongLuong;
      trongLuongXuat = typeof v === "number" && !isNaN(v) ? v : "";
    } else if (typeof tl === "number" && tl !== 0) {
      trongLuongXuat = tl;
    } else if (typeof tl === "string" && lamSach(tl) !== "") {
      trongLuongXuat = parseSoTuExcel(tl);
    }
    let theTichXuat: string | number = "";
    if (thongSoHang) {
      let v2 = thongSoHang.theTich;
      theTichXuat = typeof v2 === "number" && !isNaN(v2) ? v2 : "";
    } else if (typeof tt === "number" && tt !== 0) {
      theTichXuat = tt;
    } else if (typeof tt === "string" && lamSach(tt) !== "") {
      theTichXuat = parseSoTuExcel(tt);
    }

    let ngayNhapMoiNhat = layNgayNhapMoiNhatTuCacCot(
      valNgayNhaps.map((dataNgay) => dataNgay[i][0]),
    );

    // Gom tạm để xử lý quy tắc "xuất theo cặp" theo nhóm Hợp đồng + SKU.
    ketQuaTam.push({
      dong: [
        0, // A: STT (Sẽ điền sau khi sắp xếp)
        hd, // B: Mã hợp đồng
        mh, // C: Mã hàng
        th, // D: Tên hàng
        kichThuocXuat, // E: Kích thước
        sldh, // F: Số lượng đơn hàng
        "", // G: ĐVT (Bỏ trống)
        daXuatTamTinh, // H: Đã xuất hàng (Thành phẩm + pending nhóm)
        dongXuatDi, // I: Số lượng xuất đi
        "", // J: ĐVT (lặp lại, Bỏ trống)
        "", // K: Số lô (Bỏ trống)
        trongLuongXuat, // L: Trọng lượng 1 cái
        "", // M: Tổng trọng lượng (set formula sau khi sort/đánh STT)
        theTichXuat, // N: Thể tích 1 cái
        "", // O: Tổng thể tích (set formula sau khi sort/đánh STT)
        vt, // P: Ghi chú (Vị trí)
        "", // Q: Ghi chú mở rộng (gộp header với P)
        "", // R: Hàng rớt
        "", // S: Phiếu / ngày soạn gần nhất
        tinhMaPhanLoaiMitsuha(hd, CONFIG.uuTienTheoMaPhanLoai), // T: Mã phân loại
        tonKhoThucTe, // U: Tồn kho thực tế hiện tại
        chuaHoanThanhAmSo(chuaHoanThanh), // V: Chưa hoàn thành
        "", // W: Lý do không đạt
        ngayNhapMoiNhat, // X: Ngày nhập mới nhất (Excel serial, format dd/mm/yyyy)
      ],
      khoaNhom: `${thiTruong}|${hd}|${mh}`,
      khoaPending: khoaPendingNhom,
      nguongTheoCap: nguongXuatTheoCap,
      soLuongXuat: dongXuatDi,
      biChanTheoCap: false,
      biChanCapAB: false,
      biChanNhomChuaDuNhap: nhomHopDongSKUChuaDuNhap,
      biCanhBaoVang: nhomHopDongSKUChuaDuNhap,
      khoaCapAB: thongTinCapAB.laSKUCapAB
        ? `${thiTruong}|${hd}|${thongTinCapAB.skuGoc}`
        : "",
      hauToCapAB: thongTinCapAB.hauTo,
      hopDongKey: hd,
      maHangKey: String(mh).trim(),
      conLaiTruocXuat: conLaiCuaDon,
      lyDoKhongDat: "",
      maPhanLoai: tinhMaPhanLoaiMitsuha(hd, CONFIG.uuTienTheoMaPhanLoai),
      slHangRot: 0,
    });
  }

  // Lọc theo quy tắc "xuất theo cặp":
  // Nếu thị trường có ngưỡng (VD EU=8), nhóm HD+SKU chỉ được xuất khi tổng đủ ngưỡng đó.
  let tongXuatTheoNhom: { [khoaNhom: string]: number } = {};
  for (let dong of ketQuaTam) {
    if (dong.nguongTheoCap <= 0) continue;
    tongXuatTheoNhom[dong.khoaNhom] =
      (tongXuatTheoNhom[dong.khoaNhom] || 0) + dong.soLuongXuat;
  }

  let nhomBiChanTheoCap: { [khoaNhom: string]: boolean } = {};
  for (let dong of ketQuaTam) {
    if (dong.nguongTheoCap <= 0) continue;
    let tong = tongXuatTheoNhom[dong.khoaNhom] || 0;
    if (tong < dong.nguongTheoCap) {
      dong.biChanTheoCap = true;
      dong.biCanhBaoVang = true;
      nhomBiChanTheoCap[dong.khoaNhom] = true;
    }
  }

  // Lọc theo quy tắc SKU hậu tố A/B:
  // Cùng 1 Hợp đồng + SKU gốc phải có đủ cả A và B ở trạng thái đủ điều kiện,
  // nếu thiếu 1 bên thì chặn toàn bộ nhóm để tránh xuất tách rời.
  let nhomSKUCapAB: {
    [khoa: string]: { coA: boolean; coB: boolean; chiSoDong: number[] };
  } = {};
  for (let i = 0; i < ketQuaTam.length; i++) {
    let dong = ketQuaTam[i];
    if (!dong.khoaCapAB || !dong.hauToCapAB) continue;
    if (!nhomSKUCapAB[dong.khoaCapAB]) {
      nhomSKUCapAB[dong.khoaCapAB] = { coA: false, coB: false, chiSoDong: [] };
    }
    let nhom = nhomSKUCapAB[dong.khoaCapAB];
    nhom.chiSoDong.push(i);
    if (!dong.biCanhBaoVang) {
      if (dong.hauToCapAB === "A") nhom.coA = true;
      if (dong.hauToCapAB === "B") nhom.coB = true;
    }
  }
  let soNhomBiChanCapAB = 0;
  for (let khoa in nhomSKUCapAB) {
    let nhom = nhomSKUCapAB[khoa];
    if (nhom.coA && nhom.coB) continue;
    soNhomBiChanCapAB++;
    for (let chiSo of nhom.chiSoDong) {
      ketQuaTam[chiSo].biChanCapAB = true;
      ketQuaTam[chiSo].biCanhBaoVang = true;
    }
  }

  let layLyDoKhongDat = (x: {
    biChanNhomChuaDuNhap: boolean;
    biChanTheoCap: boolean;
    biChanCapAB: boolean;
    biCanhBaoVang: boolean;
  }): string => {
    if (x.biChanNhomChuaDuNhap) return "Nhóm HĐ+SKU còn dòng chưa nhập đủ";
    if (x.biChanTheoCap) return "Chưa đạt ngưỡng xuất theo cặp thị trường";
    if (x.biChanCapAB) return "Thiếu cặp SKU A/B";
    if (x.biCanhBaoVang) return "Cảnh báo khác / chưa đạt điều kiện";
    return "";
  };
  for (let dong of ketQuaTam) {
    dong.lyDoKhongDat = dong.lyDoKhongDat || layLyDoKhongDat(dong);
    dong.dong[22] = dong.lyDoKhongDat;
  }

  let nhanHangRot = String(cfgHangRot?.nhanViTri || "RỚT").trim() || "RỚT";
  let soDongGanNhanHangRot = 0;
  if (cfgHangRot?.bat && soNhomHangRot > 0) {
    let slRotConLai: { [khoa: string]: number } = {};
    for (let khoa in slHangRotTheoNhom) {
      slRotConLai[khoa] = slHangRotTheoNhom[khoa];
    }
    for (let item of ketQuaTam) {
      let khoa = taoKhoaHangRot(item.dong[1], item.dong[2]);
      let rot = slRotConLai[khoa] || 0;
      if (rot <= 1e-6) continue;
      let slDong = Number(item.soLuongXuat || 0);
      let gan = Math.min(rot, slDong);
      if (gan <= 1e-6) continue;
      slRotConLai[khoa] = Number((rot - gan).toFixed(4));
      item.slHangRot = gan;
      item.dong[17] = ganNhanCotHangRot(nhanHangRot, gan, slDong);
      let tt = thongTinHangRotTheoNhom[khoa];
      item.dong[18] = tt
        ? ganChuoiPhieuNgayHangRot(tt.soPhieu, tt.ngayKey)
        : "";
      soDongGanNhanHangRot++;
    }
  }

  // --- SẮP XẾP VÀ ĐÁNH LẠI STT ---
  // Sắp xếp "natural" theo vị trí để đi lấy hàng đỡ vòng:
  // ví dụ A5 < A8 < A24 < B3 ... và xử lý được chuỗi ghép như "B44C72A13".
  let tachCumViTri = (viTri: string): { khu: string; so: number }[] => {
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

    // Chuỗi ghép có thể viết không theo thứ tự (VD: B44C72A13),
    // nên chuẩn hóa về thứ tự khu/số tăng dần trước khi so sánh.
    ketQua.sort((a, b) => {
      if (a.khu !== b.khu) return a.khu.localeCompare(b.khu);
      return a.so - b.so;
    });
    return ketQua;
  };

  // So sánh 2 chuỗi vị trí theo từng cụm khu+số.
  let soSanhChuoiViTri = (vtA: string, vtB: string): number => {
    let cumA = tachCumViTri(vtA);
    let cumB = tachCumViTri(vtB);

    // Dòng có vị trí hợp lệ ưu tiên đứng trước.
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

    return vtA.localeCompare(vtB);
  };

  // So sánh 2 dòng theo vị trí ở cột P.
  let soSanhTheoViTri = (
    a: (string | number | boolean)[],
    b: (string | number | boolean)[],
  ): number => {
    let vtA = String(a[15]).trim();
    let vtB = String(b[15]).trim();
    return soSanhChuoiViTri(vtA, vtB);
  };

  // Hợp đồng được coi là "hoàn tất trong wave" nếu với mọi SKU có trong danh sách đủ điều kiện,
  // tổng SL xuất đi = tổng phần còn lại (theo từng nhóm HĐ+SKU). Chỉ đánh giá trên các dòng đang soạn.
  let tinhTapHopDongHoanTatTrongWave = (
    items: {
      hopDongKey: string;
      maHangKey: string;
      conLaiTruocXuat: number;
      soLuongXuat: number;
    }[],
  ): Set<string> => {
    let epsilon = 1e-3;
    let agg = new Map<string, { conLai: number; xuat: number }>();
    let skusTheoHd = new Map<string, Set<string>>();

    for (let it of items) {
      let k = `${it.hopDongKey}\x1f${it.maHangKey}`;
      let cur = agg.get(k);
      if (!cur) {
        cur = { conLai: 0, xuat: 0 };
        agg.set(k, cur);
      }
      cur.conLai += Number(Number(it.conLaiTruocXuat).toFixed(4));
      cur.xuat += Number(Number(it.soLuongXuat).toFixed(4));

      if (!skusTheoHd.has(it.hopDongKey)) {
        skusTheoHd.set(it.hopDongKey, new Set());
      }
      skusTheoHd.get(it.hopDongKey)!.add(it.maHangKey);
    }

    let hoanTat = new Set<string>();
    for (let hd of Array.from(skusTheoHd.keys())) {
      let ok = true;
      for (let mh of Array.from(skusTheoHd.get(hd)!)) {
        let a = agg.get(`${hd}\x1f${mh}`);
        if (!a) {
          ok = false;
          break;
        }
        let du = a.conLai - a.xuat;
        if (Math.abs(du) > epsilon) {
          ok = false;
          break;
        }
      }
      if (ok) hoanTat.add(hd);
    }
    return hoanTat;
  };

  let khoaSapXepGomTrungSKU = (dong: (string | number | boolean)[]): string => {
    let s6 = chuanHoaSKU6So(dong[2]);
    return s6 || lamSach(dong[2]);
  };

  // Dòng đủ điều kiện xuất xếp trước, dòng chưa đủ cặp đưa xuống cuối.
  let dongDuDieuKienItems = ketQuaTam.filter((x) => !x.biCanhBaoVang);
  let tapNhomHoanTatCoPending = tinhTapNhomHoanTatCoPending(
    dongDuDieuKienItems.map((x) => ({
      khoaNhom: x.khoaPending,
      conLaiTruocXuat: x.conLaiTruocXuat,
      soLuongXuat: x.soLuongXuat,
    })),
    pendingTongBanDauTheoNhom,
  );
  let apDungUuTienHdHoanTatTrongWave =
    CONFIG.uuTienHopDongHoanTatTrongWave &&
    CONFIG.uuTienHopDongHoanTatTrongWaveTrongCungViTri;
  let tapHdHoanTatTrongWave = apDungUuTienHdHoanTatTrongWave
    ? tinhTapHopDongHoanTatTrongWave(dongDuDieuKienItems)
    : new Set<string>();

  dongDuDieuKienItems.sort((a, b) => {
    if (cfgHangRot?.duaLenDauDanhSach) {
      let aRot = (a.slHangRot || 0) > 1e-6;
      let bRot = (b.slHangRot || 0) > 1e-6;
      if (aRot !== bRot) return aRot ? -1 : 1;
    }
    // Ưu tiên theo mã phân loại (S → T → M → C → F → KHAC) nếu bật
    if (dangBatUuTienPhanLoai && a.maPhanLoai !== b.maPhanLoai) {
      return soSanhMaPhanLoai(
        a.maPhanLoai,
        b.maPhanLoai,
        CONFIG.uuTienTheoMaPhanLoai,
      );
    }
    let theoViTri = soSanhTheoViTri(a.dong, b.dong);
    if (theoViTri !== 0) return theoViTri;
    if (apDungUuTienHdHoanTatTrongWave) {
      let aHT = tapHdHoanTatTrongWave.has(a.hopDongKey);
      let bHT = tapHdHoanTatTrongWave.has(b.hopDongKey);
      if (aHT !== bHT) return aHT ? -1 : 1;
    }
    if (CONFIG.gomHangCungSKU6LienKe) {
      let ka = khoaSapXepGomTrungSKU(a.dong);
      let kb = khoaSapXepGomTrungSKU(b.dong);
      if (ka !== kb) return ka.localeCompare(kb);
    }
    return 0;
  });
  let dongDuDieuKien = dongDuDieuKienItems.map((x) => x.dong);
  let dongBiChanTheoCap = ketQuaTam
    .filter((x) => x.biCanhBaoVang)
    .map((x) => x.dong)
    .sort((a, b) => {
      if (cfgHangRot?.duaLenDauDanhSach) {
        let aRot = String(a[17] || "").indexOf(nhanHangRot) === 0;
        let bRot = String(b[17] || "").indexOf(nhanHangRot) === 0;
        if (aRot !== bRot) return aRot ? -1 : 1;
      }
      // Ưu tiên theo mã phân loại cho phần cảnh báo vàng
      if (dangBatUuTienPhanLoai) {
        let maPLA = String(a[19] || "");
        let maPLB = String(b[19] || "");
        if (maPLA !== maPLB) {
          return soSanhMaPhanLoai(maPLA, maPLB, CONFIG.uuTienTheoMaPhanLoai);
        }
      }
      let theoViTri = soSanhTheoViTri(a, b);
      if (theoViTri !== 0) return theoViTri;
      if (CONFIG.gomHangCungSKU6LienKe) {
        let ka = khoaSapXepGomTrungSKU(a);
        let kb = khoaSapXepGomTrungSKU(b);
        if (ka !== kb) return ka.localeCompare(kb);
      }
      return 0;
    });

  rowsKhongDatDieuKien = ketQuaTam
    .filter((x) => x.biCanhBaoVang)
    .map((x) => {
      let dong = x.dong;
      let slXuat = parseSoTuExcel(dong[8]);
      let trongLuong1Cai = parseSoTuExcel(dong[11]);
      let theTich1Cai = parseSoTuExcel(dong[13]);
      let tongTrongLuong = Number((trongLuong1Cai * slXuat).toFixed(3));
      let tongTheTich = Number((theTich1Cai * slXuat).toFixed(3));
      let lyDo = x.lyDoKhongDat || layLyDoKhongDat(x);
      return [
        0, // STT
        dong[1], // Hợp đồng
        dong[2], // Mã hàng
        dong[3], // Tên hàng
        dong[5], // SL đơn hàng
        dong[7], // Đã xuất
        dong[8], // Đề xuất xuất
        dong[15], // Vị trí
        dong[20], // Tồn kho thực tế
        dong[21], // Chưa hoàn thành
        dong[11], // Trọng lượng 1 cái
        tongTrongLuong, // Tổng trọng lượng
        dong[13], // Thể tích 1 cái
        tongTheTich, // Tổng thể tích
        lyDo, // Lý do không đạt
      ];
    });

  // Giới hạn chỉ áp cho danh sách đủ điều kiện xuất.
  if (
    CONFIG.gioiHanMaHang > 0 &&
    dongDuDieuKien.length > CONFIG.gioiHanMaHang
  ) {
    dongDuDieuKien = dongDuDieuKien.slice(0, CONFIG.gioiHanMaHang);
  }

  let chuyenSo = (giaTri: string | number | boolean): number => {
    if (typeof giaTri === "number") return giaTri;
    let so = Number(lamSach(giaTri));
    return isNaN(so) ? 0 : so;
  };

  ketQuaData = dongDuDieuKien.concat(dongBiChanTheoCap);

  if (CONFIG.sheetKhongDatDieuKienXuat?.bat !== false) {
    let sheetKhongDat = workbook.getWorksheet(
      CONFIG.sheetKhongDatDieuKienXuat.tenSheet,
    );
    if (!sheetKhongDat) {
      sheetKhongDat = workbook.addWorksheet(
        CONFIG.sheetKhongDatDieuKienXuat.tenSheet,
      );
    }

    let usedRangeKhongDat = sheetKhongDat.getUsedRange();
    if (usedRangeKhongDat) {
      usedRangeKhongDat.clear(ExcelScript.ClearApplyTo.all);
    }

    let rowsOut: (string | number | boolean)[][] = [];
    rowsOut.push(["DANH SÁCH KHÔNG ĐẠT ĐIỀU KIỆN XUẤT"]);
    rowsOut.push([""]);
    rowsOut.push([
      "STT",
      "Hợp đồng",
      "Mã hàng",
      "Tên hàng",
      "SL đơn hàng",
      "Đã xuất",
      "Đề xuất xuất",
      "Vị trí",
      "Tồn kho thực tế",
      "Chưa hoàn thành",
      "Trọng lượng",
      "Tổng trọng lượng",
      "Thể tích",
      "Tổng thể tích",
      "Lý do không đạt",
    ]);

    let tongTrongLuongKhongDat = 0;
    let tongTheTichKhongDat = 0;
    for (let i = 0; i < rowsKhongDatDieuKien.length; i++) {
      rowsKhongDatDieuKien[i][0] = i + 1;
      tongTrongLuongKhongDat += chuyenSo(rowsKhongDatDieuKien[i][11]);
      tongTheTichKhongDat += chuyenSo(rowsKhongDatDieuKien[i][13]);
    }
    if (rowsKhongDatDieuKien.length === 0) {
      rowsOut.push([
        "",
        "",
        "",
        "Không có mã không đạt điều kiện xuất trong lần chạy này",
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
        "",
      ]);
    }
    rowsOut = rowsOut.concat(rowsKhongDatDieuKien);
    rowsOut.push([
      "",
      "",
      "",
      "TỔNG",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      Number(tongTrongLuongKhongDat.toFixed(3)),
      "",
      Number(tongTheTichKhongDat.toFixed(3)),
      "",
    ]);

    let maxColsKhongDat = 15;
    let rowsOutPadded = rowsOut.map((r) => {
      let x = r.slice();
      while (x.length < maxColsKhongDat) x.push("");
      return x;
    });
    let endRowKhongDat = rowsOutPadded.length;
    sheetKhongDat.getRange(`A1:O${endRowKhongDat}`).setValues(rowsOutPadded);

    sheetKhongDat.getRange("A1:O1").merge(false);
    sheetKhongDat.getRange("A1").getFormat().getFont().setBold(true);
    sheetKhongDat.getRange("A1").getFormat().getFont().setSize(13);
    sheetKhongDat
      .getRange("A1")
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    sheetKhongDat.getRange("A3:O3").getFormat().getFont().setBold(true);
    sheetKhongDat.getRange("A3:O3").getFormat().getFill().setColor("#FCE4D6");

    let rowTongKhongDat = endRowKhongDat;
    sheetKhongDat
      .getRange(`A${rowTongKhongDat}:O${rowTongKhongDat}`)
      .getFormat()
      .getFont()
      .setBold(true);

    if (rowsKhongDatDieuKien.length > 0) {
      let startDataRow = 4;
      let endDataRow = rowTongKhongDat - 1;
      sheetKhongDat
        .getRange(`E${startDataRow}:J${endDataRow}`)
        .setNumberFormat("0");
      sheetKhongDat
        .getRange(`K${startDataRow}:N${endDataRow}`)
        .setNumberFormat("0.000");
    }
    sheetKhongDat
      .getRange(`K${rowTongKhongDat}:N${rowTongKhongDat}`)
      .setNumberFormat("0.000");

    let fullRangeKhongDat = sheetKhongDat.getRange(`A1:O${endRowKhongDat}`);
    [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight,
      ExcelScript.BorderIndex.insideHorizontal,
      ExcelScript.BorderIndex.insideVertical,
    ].forEach((bIndex) => {
      fullRangeKhongDat
        .getFormat()
        .getRangeBorder(bIndex)
        .setStyle(ExcelScript.BorderLineStyle.continuous);
    });
    sheetKhongDat.getRange("A:O").getFormat().autofitColumns();
  }

  let tongPendingChuaApDung = 0;
  for (let khoa in pendingTheoNhom) {
    tongPendingChuaApDung += pendingTheoNhom[khoa] || 0;
  }
  for (let khoa in pendingTheoViTri) {
    tongPendingChuaApDung += pendingTheoViTri[khoa] || 0;
  }
  tongPendingChuaApDung = Number(tongPendingChuaApDung.toFixed(4));

  // Đánh STT
  let stt = 1;
  for (let j = 0; j < ketQuaData.length; j++) {
    ketQuaData[j][0] = stt; // STT
    stt++;
    let soDongSheet = COT_DICH.dongBatDau + j;
    ketQuaData[j][12] = `=L${soDongSheet}*I${soDongSheet}`; // M: Tổng trọng lượng
    ketQuaData[j][14] = `=N${soDongSheet}*I${soDongSheet}`; // O: Tổng thể tích
  }

  // ==========================================
  // 4. GHI DỮ LIỆU RA SHEET PRE-PICKING-LIST
  // ==========================================
  let tenSheetDich = taoTenSheetPrePickingList(CONFIG.locThiTruong || []);
  let targetSheet = workbook.getWorksheet(tenSheetDich);
  if (!targetSheet) {
    targetSheet = workbook.addWorksheet(tenSheetDich);
  }

  let tieuDePrePickingList: (string | number | boolean)[] = [
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
    "Phiếu / ngày soạn gần nhất",
    "Mã phân loại",
    "Tồn kho thực tế hiện tại",
    "Chưa hoàn thành",
    "Lý do không đạt",
    "Ngày nhập mới nhất",
  ];
  targetSheet.getRange("P1:Q1").unmerge();
  targetSheet
    .getRange(`A1:${COT_CUOI_AUTO}1`)
    .setValues([tieuDePrePickingList]);
  targetSheet.getRange("P1:Q1").merge(false);
  targetSheet.getRange("T:T").setColumnHidden(false);
  let fmtTieuDe = targetSheet.getRange(`A1:${COT_CUOI_AUTO}1`).getFormat();
  fmtTieuDe.getFont().setBold(true);
  fmtTieuDe.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  fmtTieuDe.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((bIndex) => {
    fmtTieuDe
      .getRangeBorder(bIndex)
      .setStyle(ExcelScript.BorderLineStyle.continuous);
  });

  // Xóa dữ liệu cũ nếu cấu hình (nội dung + định dạng; giữ lastRow trước khi xóa để dọn dòng thừa sau khi ghi ít hơn).
  let lastRowCuTruocKhiGhi = 0;
  if (CONFIG.xoaDuLieuCu) {
    lastRowCuTruocKhiGhi =
      targetSheet
        .getRange("A1048576")
        .getRangeEdge(ExcelScript.KeyboardDirection.up)
        .getRowIndex() + 1;
    if (lastRowCuTruocKhiGhi >= COT_DICH.dongBatDau) {
      let rangeCu = targetSheet.getRange(
        `A${COT_DICH.dongBatDau}:${COT_CUOI_AUTO}${lastRowCuTruocKhiGhi}`,
      );
      rangeCu.clear(ExcelScript.ClearApplyTo.all);
      datLaiDinhDangVungAuto(
        targetSheet,
        `A${COT_DICH.dongBatDau}:${COT_CUOI_AUTO}${lastRowCuTruocKhiGhi}`,
      );
    }
  }

  // Ghi dữ liệu
  if (ketQuaData.length > 0) {
    let endRow = COT_DICH.dongBatDau + ketQuaData.length - 1;
    let demDongAutoHoanTatCoPending = 0;
    let demDongPendingHoanTatCoPending = 0;
    let demDongAutoSKUTrungPending = 0;
    let demDongPendingSKUTrungAuto = 0;

    let khoaPendingTheoChiSoDongKetQua: (string | undefined)[] = new Array(
      ketQuaData.length,
    );
    let tapKhoaPendingTuAuto = new Set<string>();
    for (let i = 0; i < ketQuaData.length; i++) {
      let dong = ketQuaData[i];
      let khoa = taoKhoaPendingTheoNhom(dong[1], dong[2], dong[5]);
      khoaPendingTheoChiSoDongKetQua[i] = khoa;
      tapKhoaPendingTuAuto.add(khoa);
    }

    // Cột H giữ daXuatTamTinh (max Thành phẩm / pending) — không ghi đè bằng riêng tổng pending.

    // Ghi toàn bộ dữ liệu thuần (Value)
    let targetRange = targetSheet.getRange(
      `A${COT_DICH.dongBatDau}:${COT_CUOI_AUTO}${endRow}`,
    );
    targetRange.setValues(ketQuaData);

    // Reset định dạng toàn vùng vừa ghi (tránh giữ nền/in đậm từ lần chạy trước trên cùng ô).
    datLaiDinhDangVungAuto(
      targetSheet,
      `A${COT_DICH.dongBatDau}:${COT_CUOI_AUTO}${endRow}`,
    );

    // Dòng cũ ngoài phạm vi ghi mới.
    if (CONFIG.xoaDuLieuCu && lastRowCuTruocKhiGhi > endRow) {
      let rangeThua = `A${endRow + 1}:${COT_CUOI_AUTO}${lastRowCuTruocKhiGhi}`;
      targetSheet.getRange(rangeThua).clear(ExcelScript.ClearApplyTo.all);
      datLaiDinhDangVungAuto(targetSheet, rangeThua);
    }

    targetSheet
      .getRange(`L${COT_DICH.dongBatDau}:L${endRow}`)
      .setNumberFormat("General");
    targetSheet
      .getRange(`M${COT_DICH.dongBatDau}:M${endRow}`)
      .setNumberFormat("General");
    targetSheet
      .getRange(`N${COT_DICH.dongBatDau}:N${endRow}`)
      .setNumberFormat("0.000");
    targetSheet
      .getRange(`O${COT_DICH.dongBatDau}:O${endRow}`)
      .setNumberFormat("0.000");
    targetSheet
      .getRange(`W${COT_DICH.dongBatDau}:W${endRow}`)
      .getFormat()
      .setWrapText(true);
    targetSheet
      .getRange(
        `${COT_DICH.ngayNhapMoiNhat || "X"}${COT_DICH.dongBatDau}:${COT_DICH.ngayNhapMoiNhat || "X"}${endRow}`,
      )
      .setNumberFormat("dd/mm/yyyy");
    targetSheet
      .getRange(
        `${COT_DICH.tonKhoThucTe || "U"}${COT_DICH.dongBatDau}:${COT_DICH.tonKhoThucTe || "U"}${endRow}`,
      )
      .setNumberFormat("0");
    targetSheet
      .getRange(
        `${COT_DICH.chuaHoanThanh || "V"}${COT_DICH.dongBatDau}:${COT_DICH.chuaHoanThanh || "V"}${endRow}`,
      )
      .setNumberFormat("0");

    // Nền vàng: phần cuối = chưa đạt chuẩn xuất (thiếu cặp / cảnh báo), tách khỏi phần đủ điều kiện.
    if (dongBiChanTheoCap.length > 0) {
      let startRowBiChan = COT_DICH.dongBatDau + dongDuDieuKien.length;
      let endRowBiChan = startRowBiChan + dongBiChanTheoCap.length - 1;
      targetSheet
        .getRange(`A${startRowBiChan}:${COT_CUOI_AUTO}${endRowBiChan}`)
        .getFormat()
        .getFill()
        .setColor(CONFIG.maNenCanhBaoKhongDatChuan);
    }

    // Tô màu các dòng trong auto mà pending + wave hiện tại vừa đủ hoàn tất đơn hàng.
    if (tapNhomHoanTatCoPending.size > 0 && CONFIG.maMauHoanTatDonCoPending) {
      for (let i = 0; i < dongDuDieuKien.length; i++) {
        let dong = dongDuDieuKien[i];
        let khoaPending = khoaPendingTheoChiSoDongKetQua[i];
        if (!khoaPending || !tapNhomHoanTatCoPending.has(khoaPending)) continue;

        let soDongSheet = COT_DICH.dongBatDau + i;
        targetSheet
          .getRange(`A${soDongSheet}:${COT_CUOI_AUTO}${soDongSheet}`)
          .getFormat()
          .getFill()
          .setColor(CONFIG.maMauHoanTatDonCoPending);
        demDongAutoHoanTatCoPending++;
      }
    }

    if (
      suaDinhDangSheetPending &&
      pendingSheet &&
      pendingLastRow >= COT_PENDING.dongBatDau
    ) {
      pendingSheet
        .getRange(`A${COT_PENDING.dongBatDau}:P${pendingLastRow}`)
        .getFormat()
        .getFill()
        .clear();

      if (tapNhomHoanTatCoPending.size > 0 && CONFIG.maMauHoanTatDonCoPending) {
        for (let meta of pendingRowsMeta) {
          if (!tapNhomHoanTatCoPending.has(meta.khoaPending)) continue;
          pendingSheet
            .getRange(`A${meta.soDongSheet}:P${meta.soDongSheet}`)
            .getFormat()
            .getFill()
            .setColor(CONFIG.maMauHoanTatDonCoPending);
          demDongPendingHoanTatCoPending++;
        }
      }

      if (
        CONFIG.toMauSKUAutoTrungPending &&
        CONFIG.maMauNenSKUTrungPending &&
        tapKhoaPendingTuAuto.size > 0
      ) {
        for (let meta of pendingRowsMeta) {
          if (!meta.khoaPending || !tapKhoaPendingTuAuto.has(meta.khoaPending))
            continue;
          let fmtP = pendingSheet
            .getRange(`${cotToMauSKUPending}${meta.soDongSheet}`)
            .getFormat();
          fmtP.getFill().setColor(CONFIG.maMauNenSKUTrungPending);
          if (CONFIG.maMauChuSKUTrungPending) {
            fmtP.getFont().setColor(CONFIG.maMauChuSKUTrungPending);
          }
          demDongPendingSKUTrungAuto++;
        }
      }

      pendingSheet.getRange("A:P").getFormat().autofitColumns();
    }

    // Tô màu các SKU trùng (6–7 số): chỉ ô cột C; áp cả phần đủ điều kiện và phần nền vàng (ô C ghi đè nền vàng).
    if (CONFIG.toMauSKUTrung && CONFIG.maMauSKUTrung) {
      let soLanXuatHienSKU: { [sku: string]: number } = {};
      for (let dong of ketQuaData) {
        let sku = chuanHoaSKU6So(dong[2]);
        if (!sku) continue;
        soLanXuatHienSKU[sku] = (soLanXuatHienSKU[sku] || 0) + 1;
      }

      let demDongSKUTrung = 0;
      for (let i = 0; i < ketQuaData.length; i++) {
        let sku = chuanHoaSKU6So(ketQuaData[i][2]);
        if (!sku || (soLanXuatHienSKU[sku] || 0) < 2) continue;

        let soDongSheet = COT_DICH.dongBatDau + i;

        targetSheet
          .getRange(`C${soDongSheet}`)
          .getFormat()
          .getFill()
          .setColor(CONFIG.maMauSKUTrung);
        demDongSKUTrung++;
      }
      if (demDongSKUTrung > 0) {
        console.log(
          `• Đã tô màu ${demDongSKUTrung} ô cột SKU bị trùng (${CONFIG.maMauSKUTrung}).`,
        );
      }
    }

    if (
      CONFIG.toMauSKUAutoTrungPending &&
      CONFIG.maMauNenSKUTrungPending &&
      tapKhoaPendingTuPending.size > 0
    ) {
      for (let i = 0; i < ketQuaData.length; i++) {
        let dong = ketQuaData[i];
        let khoa = khoaPendingTheoChiSoDongKetQua[i];
        if (!khoa || !tapKhoaPendingTuPending.has(khoa)) continue;
        let soDongSheet = COT_DICH.dongBatDau + i;
        let fmt = targetSheet.getRange(`C${soDongSheet}`).getFormat();
        fmt.getFill().setColor(CONFIG.maMauNenSKUTrungPending);
        if (CONFIG.maMauChuSKUTrungPending) {
          fmt.getFont().setColor(CONFIG.maMauChuSKUTrungPending);
        }
        demDongAutoSKUTrungPending++;
      }
      if (demDongAutoSKUTrungPending > 0 || demDongPendingSKUTrungAuto > 0) {
        console.log(
          `• Dòng trùng khóa pending (HĐ+mã+SL đơn) giữa "${tenSheetDich}" và "${COT_PENDING.tenSheet}": đã tô ${demDongAutoSKUTrungPending} ô cột C (pre-picking), ${demDongPendingSKUTrungAuto} ô cột C (pending) — nền ${CONFIG.maMauNenSKUTrungPending}, chữ ${CONFIG.maMauChuSKUTrungPending || "mặc định"}.`,
        );
      }
    }
    if (CONFIG.traCuuTenMo?.maMauNenKichThuoc) {
      let demDongKichThuocTraCuuMo = 0;
      for (let i = 0; i < ketQuaData.length; i++) {
        let dong = ketQuaData[i];
        if (
          !laKichThuocLayTuTraCuuTenMo(
            lamSach(dong[2]),
            chuanHoaTenSanPham(dong[3]),
            dong[4],
            lookupThongSo,
            lookupThongSoTheoTen,
            danhSachThongSoTheoTen,
            CONFIG.traCuuTenMo,
          )
        ) {
          continue;
        }
        let soDongSheet = COT_DICH.dongBatDau + i;
        targetSheet
          .getRange(`E${soDongSheet}`)
          .getFormat()
          .getFill()
          .setColor(CONFIG.traCuuTenMo.maMauNenKichThuoc);
        demDongKichThuocTraCuuMo++;
      }
      if (demDongKichThuocTraCuuMo > 0) {
        console.log(
          `• Đã tô nền ${demDongKichThuocTraCuuMo} ô cột E (kích thước từ tra cứu mờ theo tên — ${CONFIG.traCuuTenMo.maMauNenKichThuoc}).`,
        );
      }
    }

    if (demDongAutoHoanTatCoPending > 0) {
      console.log(
        `• Đã tô màu ${demDongAutoHoanTatCoPending} dòng trong "${tenSheetDich}" thuộc nhóm pending được hoàn tất đơn hàng (${CONFIG.maMauHoanTatDonCoPending}).`,
      );
    }
    if (demDongPendingHoanTatCoPending > 0) {
      console.log(
        `• Đã tô màu ${demDongPendingHoanTatCoPending} dòng trong "${COT_PENDING.tenSheet}" thuộc nhóm pending được hoàn tất đơn hàng (${CONFIG.maMauHoanTatDonCoPending}).`,
      );
    }
    // Căn chỉnh tự động kích thước cột để dễ nhìn (bỏ qua cột quá rộng)
    targetSheet.getRange(`A:${COT_CUOI_AUTO}`).getFormat().autofitColumns();

    console.log(`\n=========================================`);
    console.log(`     HOÀN TẤT SOẠN HÀNG XUẤT`);
    console.log(`=========================================`);
    console.log(
      `• Đã tìm thấy ${ketQuaData.length} dòng đủ điều kiện xuất kho (áp dụng ngưỡng theo thị trường hoặc đã xong đơn).`,
    );
    if (cfgHangRot?.bat) {
      console.log(
        `• Hàng rớt: ${soNhomHangRot} nhóm | tổng SL ${tongSlHangRot} | ${soDongGanNhanHangRot} dòng gắn nhãn "${nhanHangRot}" (lấy ở khu tập kết, không về vị trí kho cũ).`,
      );
    }
    if (boQuaPendingViKhongCoSheet) {
      console.log(
        `• Không tìm thấy sheet "${COT_PENDING.tenSheet}", đang bỏ qua dữ liệu pending.`,
      );
    }
    if (soDongPendingHopLe > 0) {
      console.log(
        `• Đã đọc ${soDongPendingHopLe} dòng pending từ sheet "${COT_PENDING.tenSheet}" (tổng SL ${tongPendingTrongSheet}).`,
      );
    }
    if (tongPendingDaApDung > 0) {
      console.log(
        `• Đã trừ tạm ${tongPendingDaApDung} khỏi tồn khả dụng trên ${soDongNguonDaTruPending} dòng nguồn.`,
      );
    }
    if (tongPendingChuaApDung > 0) {
      console.log(
        `• Còn ${tongPendingChuaApDung} trong sheet "${COT_PENDING.tenSheet}" chưa khớp dữ liệu nguồn hiện tại (kiểm tra HĐ/Mã/SL đơn/Vị trí).`,
      );
    }
    if (apDungUuTienHdHoanTatTrongWave) {
      console.log(
        `• Ưu tiên HĐ hoàn tất trong wave (trong cùng ô vị trí): ${tapHdHoanTatTrongWave.size} hợp đồng.`,
      );
    } else if (CONFIG.uuTienHopDongHoanTatTrongWave) {
      console.log(
        `• Cùng ô vị trí: xếp theo SKU — quét sạch line trước khi sang ô kế (uuTienHopDongHoanTatTrongWaveTrongCungViTri = false).`,
      );
    }
    let soNhomBiChan = Object.keys(nhomBiChanTheoCap).length;
    let tapNhomBiChanChuaDuNhap = new Set<string>();
    for (let dong of ketQuaTam) {
      if (dong.biChanNhomChuaDuNhap) tapNhomBiChanChuaDuNhap.add(dong.khoaNhom);
    }
    if (soNhomBiChan > 0) {
      console.log(
        `• Có ${soNhomBiChan} nhóm Hợp đồng+SKU chưa đủ ngưỡng xuất theo cặp: đã đưa xuống cuối danh sách và tô vàng.`,
      );
    }
    if (tapNhomBiChanChuaDuNhap.size > 0) {
      console.log(
        `• Có ${tapNhomBiChanChuaDuNhap.size} nhóm Hợp đồng+SKU còn dòng chưa nhập đủ: đã chặn khỏi phần soạn và tô vàng.`,
      );
    }
    if (soNhomBiChanCapAB > 0) {
      console.log(
        `• Có ${soNhomBiChanCapAB} nhóm SKU hậu tố A/B chưa đủ cặp A+B: đã đưa xuống cuối danh sách và tô vàng.`,
      );
    }
    if (dongBiChanTheoCap.length > 0) {
      console.log(
        `• Có ${dongBiChanTheoCap.length} dòng chưa đạt chuẩn xuất (thiếu cặp / cảnh báo): đã đưa xuống cuối danh sách và tô nền vàng (${CONFIG.maNenCanhBaoKhongDatChuan}).`,
      );
    }
    if (tapSKUThieuThongSo.size > 0) {
      let dsSKUThieu = Array.from(tapSKUThieuThongSo).slice(0, 20);
      console.log(
        `• SKU thiếu thông số trong sheet "${CONFIG.tenSheetThongSo}": ${tapSKUThieuThongSo.size} mã (${dsSKUThieu.join(", ")}${tapSKUThieuThongSo.size > dsSKUThieu.length ? ", ..." : ""}).`,
      );
    }
    if (dangBatSkuKhongSoan) {
      console.log(
        `• Đã bỏ qua ${soDongBoQuaSkuKhongSoan} dòng nguồn theo skuKhongDuocSoan (${CONFIG.skuKhongDuocSoan.join(", ")}).`,
      );
    }
    if (dangBatUuTienPhanLoai) {
      let thuTuHienThi = chuanHoaDanhSachKyTuPhanLoai(
        CONFIG.uuTienTheoMaPhanLoai.thuTuMacDinh,
      ).join(" → ");
      console.log(
        `• Ưu tiên theo mã phân loại: ${thuTuHienThi} → KHAC (đang BẬT).`,
      );
    }
    console.log(`• Dữ liệu đã được ghi vào sheet "${tenSheetDich}".`);
    console.log(`=========================================`);
  } else {
    let demDongPendingHoanTatCoPending = 0;
    if (
      suaDinhDangSheetPending &&
      pendingSheet &&
      pendingLastRow >= COT_PENDING.dongBatDau
    ) {
      pendingSheet
        .getRange(`A${COT_PENDING.dongBatDau}:P${pendingLastRow}`)
        .getFormat()
        .getFill()
        .clear();

      if (tapNhomHoanTatCoPending.size > 0 && CONFIG.maMauHoanTatDonCoPending) {
        for (let meta of pendingRowsMeta) {
          if (!tapNhomHoanTatCoPending.has(meta.khoaPending)) continue;
          pendingSheet
            .getRange(`A${meta.soDongSheet}:P${meta.soDongSheet}`)
            .getFormat()
            .getFill()
            .setColor(CONFIG.maMauHoanTatDonCoPending);
          demDongPendingHoanTatCoPending++;
        }
      }

      pendingSheet.getRange("A:P").getFormat().autofitColumns();
    }
    console.log(
      `\nKhông tìm thấy dữ liệu nào thỏa mãn điều kiện xuất kho (áp dụng ngưỡng theo thị trường hoặc đã xong đơn).`,
    );
    if (boQuaPendingViKhongCoSheet) {
      console.log(
        `• Không tìm thấy sheet "${COT_PENDING.tenSheet}", đang bỏ qua dữ liệu pending.`,
      );
    }
    if (soDongPendingHopLe > 0) {
      console.log(
        `• Đã đọc ${soDongPendingHopLe} dòng pending từ sheet "${COT_PENDING.tenSheet}" (tổng SL ${tongPendingTrongSheet}).`,
      );
    }
    if (tongPendingDaApDung > 0) {
      console.log(
        `• Đã trừ tạm ${tongPendingDaApDung} khỏi tồn khả dụng trên ${soDongNguonDaTruPending} dòng nguồn.`,
      );
    }
    if (tongPendingChuaApDung > 0) {
      console.log(
        `• Còn ${tongPendingChuaApDung} trong sheet "${COT_PENDING.tenSheet}" chưa khớp dữ liệu nguồn hiện tại (kiểm tra HĐ/Mã/SL đơn/Vị trí).`,
      );
    }
    if (demDongPendingHoanTatCoPending > 0) {
      console.log(
        `• Đã tô màu ${demDongPendingHoanTatCoPending} dòng trong "${COT_PENDING.tenSheet}" thuộc nhóm pending được hoàn tất đơn hàng (${CONFIG.maMauHoanTatDonCoPending}).`,
      );
    }
    if (dangBatLocSKU) {
      console.log(
        `Lưu ý: Bạn đang bật bộ lọc SKU: "${CONFIG.locSKU.join(", ")}"`,
      );
    }
    if (dangBatSkuKhongSoan) {
      console.log(
        `Lưu ý: Đang loại ${CONFIG.skuKhongDuocSoan.length} mục khỏi soạn (skuKhongDuocSoan); đã bỏ qua ${soDongBoQuaSkuKhongSoan} dòng nguồn.`,
      );
    }
    if (CONFIG.locThiTruong && CONFIG.locThiTruong.length > 0) {
      console.log(
        `Lưu ý: Bạn đang bật bộ lọc thị trường: "${CONFIG.locThiTruong.join(", ")}"`,
      );
    }
    if (CONFIG.locTonKho) {
      console.log(
        `Lưu ý: Ngưỡng tồn kho đang áp dụng theo thị trường. Mặc định: ${CONFIG.tonKhoToiThieu}; Ghi đè: ${Object.entries(
          CONFIG.tonKhoTheoThiTruong,
        )
          .map(([tt, nguong]) => `${tt}=${nguong}`)
          .join(", ")}`,
      );
    }
    if (Object.keys(CONFIG.xuatTheoCapTheoThiTruong).length > 0) {
      console.log(
        `Lưu ý: Quy tắc xuất theo cặp đang bật: ${Object.entries(
          CONFIG.xuatTheoCapTheoThiTruong,
        )
          .map(([tt, nguong]) => `${tt}=${nguong}`)
          .join(", ")}`,
      );
    }
    if (dangBatLocNgayChiTiet) {
      let moTaBoLocNgay = `ngày=[${boLocNgayChiTiet.ngay.join(
        ", ",
      )}], tháng=[${boLocNgayChiTiet.thang.join(
        ", ",
      )}], năm=[${boLocNgayChiTiet.nam.join(", ")}]`;
      console.log(`Lưu ý: Bạn đang bật bộ lọc ngày nhập: ${moTaBoLocNgay}`);
    }
    if (CONFIG.gioiHanMaHang > 0) {
      console.log(
        `Lưu ý: Bạn đang giới hạn chỉ lấy tối đa ${CONFIG.gioiHanMaHang} mã hàng.`,
      );
    }
    console.log(`• Sheet "${tenSheetDich}" đã được tạo/cập nhật tiêu đề.`);
  }
}
