/**
 * ===================================================================
 *  Auto_Luffy – Tự động đổi dữ liệu từ sheet1 sang sheet ERP.
 * ===================================================================
 *
 * Script này tự động đổi dữ liệu từ sheet1 sang sheet ERP dựa trên cấu hình.
 * Mọi thông số đều nằm trong CONFIG → dễ tùy chỉnh, mở rộng.
 *
 */
// @ts-ignore: Office Scripts — tránh cảnh báo trùng tên khi gộp nhiều file TS
const LUFFY_CONFIG = {
  tenSheetNguon: "sale",
  /** Để trống "" = sheet đang mở — A:C tự nhập; I–R (gồm P,Q,R tạm + XLOOKUP kích thước theo E) + công thức còn lại */
  tenSheetDich: "ERP",
  dataTable: {
    /** Nhận diện hàng tiêu đề Sheet1 (cột A), giống Auto_PXK */
    headerKeywords: ["STT", "序号"] as string[],
  },
  /** Sheet1: cột số thứ tự (STT) — xác định vùng dữ liệu: 1,2,3… đến hàng STT trống đầu tiên */
  cotSoThuTu: "A",
  /** Sheet1: cột nguồn — Mã HĐ (B), SKU (C), Số lượng (I), Kích thước (E) */
  nguon: { hopDong: "B", sku: "C", sl: "I", kichThuoc: "E" },
  /**
   * true: sau khi chạy, đọc lại cột nhập tay trên ERP (dichVung khối duLieuERP — mặc định cột sl = C),
   * bỏ tối đa 10 số 0 cuối (phần nguyên) và ghi number. Không áp dụng cho dữ liệu lấy từ Sheet1. false: không đổi.
   */
  chuyenCotCSkuThanhSoBo10So0Cuoi: false,
  /**
   * Khối 1: A,B,C do người dùng — D=B&A; E=UNIQUE(FILTER(D…,D…<>"")) nếu uniqueEFilterBoTrong, không thì UNIQUE(D…); F=STT(E#).
   * Khối 2: L,M,N từ Sheet1; P,Q,R = raw dài/rộng/cao; I,J,K = tra theo E khớp O; O=M&L.
   */
  dichVung: [
    {
      duLieuERP: false,
      hopDong: "A",
      sku: "B",
      sl: "C",
      code: "D",
      unique: "E",
    },
    {
      duLieuDataSheet: false,
      dai: "I",
      rong: "J",
      cao: "K",
      hopDong: "L",
      sku: "M",
      sl: "N",
      code: "O",
    },
  ] as const,
  /**
   * Cột tạm raw dài/rộng/cao (cùng số dòng Sheet1) — dùng cho XLOOKUP I,J,K theo cột E (code unique).
   * Phải nằm bên phải O, không trùng vùng công thức khác.
   */
  cotTamKichThuoc: { dai: "P", rong: "Q", cao: "R" },
  /**
   * F{dongCongThucUnique}:F… = STT spill theo danh sách Code Unique ở E.
   */
  cotSttCodeUnique: "F",
  /**
   * G{bd}:G… = SUMIFS(N; O; E cùng dòng) — tổng số lượng theo code O (= M&L) trùng mã unique ở E.
   */
  sumIfsTheoMaDoc: {
    cotKetQua: "G",
    cotSum: "N",
    cotKey: "O",
    cotTieuChi: "E",
  },
  /** H{bd}:H… = C - G (số lượng nhập tay trừ tổng Sheet1 cùng mã). */
  cotSoLuongConLai: "H",
  /** Dòng đầu ghi dữ liệu trên sheet đích (luôn cố định, không bằng dòng đầu Sheet1) */
  dongBatDau: 3,
  /** Bỏ dòng Sheet1 trống (B+C+H) ở đầu/cuối vùng — tránh dữ liệu thật lệch xuống H7… */
  boQuaDongTrongNguon: true,
  /** Dòng tiêu đề cột E / F (vd. 2) — không xóa khi dọn spill */
  dongTieuDeCotE: 2,
  /** Dòng công thức spill: E3 (UNIQUE hoặc UNIQUE+FILTER) và F3 (STT) — phải > dongTieuDeCotE */
  dongCongThucUnique: 3,
  /**
   * true: E = UNIQUE(FILTER(D{bd}:D{cuối}, D{bd}:D{cuối}<>"")) — bỏ ô trống cột D trước khi unique.
   * false: E = UNIQUE(D{bd}:D{cuối}). Công thức gán dùng dấu phẩy (en-US) cho Office Scripts.
   */
  uniqueEFilterBoTrong: true,
  xoaCotEdenDong: 50000,
  /**
   * Danh sách tên sheet cần tự động tạo khi chạy script.
   * Nếu sheet đã tồn tại → bỏ qua. Nếu chưa có → tạo mới.
   * Để mảng rỗng [] nếu không muốn tự tạo sheet nào.
   */
  autoTaoSheet: ["ERP", "ERP-A"] as string[],
  /**
   * Cấu hình cấu trúc sheet ERP-A (bảng tổng hợp nhập tay + SUMIFS đối chiếu ERP).
   * - tenSheet: tên sheet ERP-A (phải nằm trong autoTaoSheet hoặc đã tồn tại).
   * - tieuDe: mảng tiêu đề hàng 1 (tương ứng cột A→G).
   * - cot: ánh xạ cột.
   * - dongTieuDe: dòng tiêu đề (1).
   * - dongBatDauDuLieu: dòng dữ liệu đầu tiên (2).
   * - soDongDuLieu: số dòng công thức được đặt sẵn (0 = theo số dòng nguồn).
   * - tenSheetErpThamChieu: tên sheet ERP dùng trong SUMIFS (mặc định lấy từ tenSheetDich).
   *   Đặt "" để tự lấy từ tenSheetDich.
   * - cotErpSumSl: cột số lượng trên ERP dùng cho SUMIFS (mặc định N).
   * - cotErpCode: cột code trên ERP dùng cho SUMIFS criteria (mặc định O).
   */
  cauHinhErpA: {
    tenSheet: "ERP-A",
    tieuDe: [
      "STT",
      "Mã hợp đồng",
      "SKU",
      "Số lượng",
      "Code",
      "Tổng",
      "Chênh lệch",
    ],
    cot: {
      stt: "A",
      hopDong: "B",
      sku: "C",
      sl: "D",
      code: "E",
      tong: "F",
      chenhLech: "G",
    },
    dongTieuDe: 1,
    dongBatDauDuLieu: 2,
    /** Số dòng công thức đặt sẵn. 0 = lấy theo số dòng nguồn (n). */
    soDongDuLieu: 0,
    /** Tên sheet ERP dùng trong SUMIFS. "" = tự lấy từ tenSheetDich. */
    tenSheetErpThamChieu: "",
    /** Cột trên ERP cho SUMIFS sum_range (mặc định N = sl từ sheet nguồn). */
    cotErpSumSl: "N",
    /** Cột trên ERP cho SUMIFS criteria_range (mặc định O = code từ sheet nguồn). */
    cotErpCode: "O",
  },
  /** Cấu hình định dạng tiêu đề (Header) và phân nhóm màu nền */
  formatHeader: {
    fontName: "Times New Roman",
    fontSize: 14,
    bold: false, // Không in đậm tiêu đề
    fontColor: "#FFFFFF", // Màu chữ trắng cho tiêu đề
    fillColorDefault: "#4F81BD", // Màu nền tiêu đề mặc định (sheet ERP-A)
    fillColorsERP: {
      nhomAF: "#2F5597", // Nhóm A-F (Mã HĐ, SKU, SL, Code, Unique, STT): Xanh dương đậm
      nhomGH: "#4F81BD", // Nhóm G-H (Tổng, Còn lại): Xanh dương nhạt
      nhomIK: "#ED7D31", // Nhóm I-K (Kích thước Dài, Rộng, Cao): Cam
      nhomLR: "#4472C4", // Nhóm L-R (Data từ Sheet1 + Cột tạm): Xanh lam
    },
    subtotalFillColor: "#FFF2CC", // Màu nền nổi bật cho các ô SUBTOTAL (G1, N1)
    subtotalFontColor: "#000000", // Màu chữ đen cho các ô SUBTOTAL (G1, N1)
  },
  /** Cấu hình công thức FILTER tại A3 trên sheet ERP để lấy dữ liệu từ sheet đối chiếu */
  congThucFilterA3ERP: {
    oBatDau: "A3",
    vungLay: "B2:D1000",
    vungKiemTra: "B2:B1000",
  },
  /** Cấu hình cảnh báo tô màu dữ liệu trùng lặp cột O (Code) trên sheet ERP */
  canhBaoTrungLapCotO: {
    bat: true,
    cot: "O",
    mauNenCanhBao: "#FFC7CE", // Nền đỏ hồng cảnh báo trùng lặp
    mauChuCanhBao: "#9C0006", // Chữ đỏ đậm cảnh báo trùng lặp
  },
  /** Mảng tiêu đề dòng 2 cho sheet ERP (cột A → R) */
  tieuDeERP: [
    "Mã hợp đồng",
    "SKU",
    "Số lượng",
    "Code",
    "Code Unique",
    "STT",
    "Tổng",
    "Còn lại",
    "Dài",
    "Rộng",
    "Cao",
    "Mã hợp đồng",
    "SKU",
    "Số lượng",
    "Code",
    "Dài",
    "Rộng",
    "Cao",
  ] as string[],
};

// End of CONFIG

function luffyLamSach(giaTri: string | number | boolean): string {
  return String(giaTri)
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();
}

/**
 * Chuỗi kích thước dạng dài*rộng*cao (vd. 113,9*58,7*13,1 hoặc 23*23*50).
 * Dấu phẩy trong từng phần được hiểu là thập phân (locale EU); ghi số khi parse được.
 */
function luffyTachKichThuoc(
  raw: string | number | boolean,
): [string | number, string | number, string | number] {
  if (raw === null || raw === undefined) return ["", "", ""];
  if (typeof raw === "number" && isFinite(raw)) {
    return ["", "", ""];
  }
  let s = luffyLamSach(String(raw));
  if (s === "") return ["", "", ""];

  let parseMotPhan = (p: string): string | number => {
    let t = p.trim();
    if (t === "") return "";
    let chuan = t.replace(/,/g, ".");
    let n = Number(chuan);
    if (!isNaN(n) && isFinite(n)) return n;
    return t;
  };

  let parts = s.split("*").map((p) => p.trim());
  let dai = parseMotPhan(parts[0] ?? "");
  let rong = parseMotPhan(parts[1] ?? "");
  let cao = parseMotPhan(parts[2] ?? "");
  return [dai, rong, cao];
}

/** Dòng nguồn “trống”: không mã HĐ, không SKU, không số lượng (số 0 vẫn tính là có SL). */
function luffyHangNguonTrong(
  hopDong: string | number | boolean,
  sku: string | number | boolean,
  soLuong: string | number | boolean,
): boolean {
  let hd = luffyLamSach(hopDong) === "";
  let ma = luffyLamSach(sku) === "";
  if (!hd || !ma) return false;
  if (typeof soLuong === "number" && !isNaN(soLuong)) return false;
  return luffyLamSach(String(soLuong)) === "";
}

/** Giống Auto_PXK: quét cột A trong usedRange, hàng chứa đủ headerKeywords (0-based). */
function luffyFindHeaderRow(
  sheet: ExcelScript.Worksheet,
  keywords: string[],
): number {
  let usedRange = sheet.getUsedRange();
  if (!usedRange) return -1;

  let lastUsedRow = usedRange.getRowCount();
  let colAValues = sheet.getRange(`A1:A${lastUsedRow}`).getValues();

  for (let i = 0; i < colAValues.length; i++) {
    let cellValue = String(colAValues[i][0]).toUpperCase();
    let allMatch = keywords.every((kw) => cellValue.includes(kw.toUpperCase()));
    if (allMatch) return i;
  }
  return -1;
}

/** Giống Auto_PXK: chữ ký từ sheet "footer" để cắt vùng quét trước footer. */
function luffyGetFooterSignatureFromWorkbook(
  workbook: ExcelScript.Workbook,
): string | null {
  let footerSheet = workbook.getWorksheet("footer");
  if (!footerSheet) return null;
  let fullUsedRange = footerSheet.getUsedRange();
  if (!fullUsedRange) return null;
  let footerTop = fullUsedRange.getIntersection("1:1");
  if (!footerTop) return null;
  let row = footerTop.getValues()[0] as unknown[];
  let parts: string[] = [];
  for (let c of row) {
    let s = String(c ?? "").trim();
    if (s) parts.push(s);
  }
  if (parts.length === 0) return null;
  let best = parts.reduce((a, b) => (a.length >= b.length ? a : b), "");
  if (best.length >= 4) return best;
  let joined = parts.join(" ");
  return joined.length >= 6 ? joined : null;
}

/** Giống Auto_PXK: dòng 1-based đầu tiên từ fromRow mà cột A chứa chữ ký footer. */
function luffyFindFooterStartRowFromSignature(
  sheet: ExcelScript.Worksheet,
  fromRow: number,
  toRow: number,
  signature: string,
): number | null {
  let needle = signature.trim().toUpperCase();
  if (needle.length < 4) return null;
  if (needle.length > 40) needle = needle.substring(0, 40);
  let colA = sheet.getRange(`A${fromRow}:A${toRow}`).getValues();
  for (let i = 0; i < colA.length; i++) {
    let a = String(colA[i][0] ?? "")
      .trim()
      .toUpperCase();
    if (a.length > 0 && a.includes(needle)) return fromRow + i;
  }
  return null;
}

/** Dòng cuối cùng được quét (footer / usedRange), giống phần giới hạn trong Auto_PXK. */
function luffyLayDongCuoiQuet(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  workbook: ExcelScript.Workbook,
): number {
  let usedRange = sheet.getUsedRange();
  if (!usedRange) return startRow;
  let totalRowsInSheet = usedRange.getRowIndex() + usedRange.getRowCount();
  if (totalRowsInSheet < startRow) return startRow;
  let scanEndRow = totalRowsInSheet;
  let sig = luffyGetFooterSignatureFromWorkbook(workbook);
  if (sig) {
    let footerStart = luffyFindFooterStartRowFromSignature(
      sheet,
      startRow,
      totalRowsInSheet,
      sig,
    );
    if (footerStart !== null && footerStart > startRow) {
      scanEndRow = footerStart - 1;
      console.log(
        `Giới hạn bảng ${sheet.getName()} trước footer (dòng footer: ${footerStart}).`,
      );
    }
  }
  return scanEndRow;
}

/** STT nguyên dương từ ô Excel (số hoặc chuỗi chỉ gồm chữ số). */
function luffyLaySoStt(raw: string | number | boolean): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number" && !isNaN(raw)) {
    let r = Math.round(raw);
    if (Math.abs(raw - r) > 1e-9) return null;
    return r >= 1 ? r : null;
  }
  let s = luffyLamSach(String(raw));
  if (!s) return null;
  if (!/^\d+$/.test(s)) return null;
  let n = parseInt(s, 10);
  return n >= 1 ? n : null;
}

/**
 * Vùng dữ liệu: từ startRow, cột STT phải là 1, 2, 3… liên tục.
 * Dừng tại dòng đầu tiên STT trống (kết thúc bảng). Sai thứ tự → ok: false + loi (gọi console.log ngoài vòng lặp).
 */
function luffyTimVungTheoCotStt(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  workbook: ExcelScript.Workbook,
  cotStt: string,
):
  | { ok: true; startRow: number; lastRow: number }
  | { ok: false; loi: string } {
  let scanEnd = luffyLayDongCuoiQuet(sheet, startRow, workbook);
  if (scanEnd < startRow) {
    return {
      ok: false,
      loi: "STT: vùng quét không hợp lệ (scanEnd < startRow).",
    };
  }
  let vals = sheet
    .getRange(`${cotStt}${startRow}:${cotStt}${scanEnd}`)
    .getValues();
  let mongDoi = 1;
  for (let i = 0; i < vals.length; i++) {
    let excelRow = startRow + i;
    let v = luffyLaySoStt(vals[i][0]);
    if (v === null) {
      if (i === 0) {
        return {
          ok: false,
          loi: `STT (${cotStt}): dòng ${excelRow} phải bắt đầu bằng 1 — ô trống hoặc không phải số thứ tự hợp lệ.`,
        };
      }
      return { ok: true, startRow, lastRow: excelRow - 1 };
    }
    if (v !== mongDoi) {
      return {
        ok: false,
        loi: `STT (${cotStt}): dòng ${excelRow} có giá trị ${v}, cần ${mongDoi} (phải liên tục 1→2→3…).`,
      };
    }
    mongDoi++;
  }
  return { ok: true, startRow, lastRow: startRow + vals.length - 1 };
}

/** Ô số lượng coi là trống (không có số hợp lệ). Số 0 vẫn là có SL. */
function luffyTrongSoLuong(sl: string | number | boolean): boolean {
  if (typeof sl === "number" && !isNaN(sl)) return false;
  return luffyLamSach(String(sl)) === "";
}

/** Dòng cuối có dữ liệu thật trong một dải cột trên sheet đích; null nếu cả vùng đều trống. */
function luffyLayDongCuoiCoDuLieuTrongVung(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  endRow: number,
  cotBatDau: string,
  cotKetThuc: string,
): number | null {
  if (endRow < startRow) return null;
  let vals = sheet
    .getRange(`${cotBatDau}${startRow}:${cotKetThuc}${endRow}`)
    .getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    let coDuLieu = vals[i].some((o) => luffyLamSach(String(o ?? "")) !== "");
    if (coDuLieu) return startRow + i;
  }
  return null;
}

/**
 * Chỉ dùng khi bật cấu hình: chuẩn hóa ô nhập tay trên sheet đích (mặc định cột C = sl khối không lấy Sheet1).
 * Bỏ tối đa 10 chữ số 0 ở cuối (phần nguyên), rồi ghi number.
 * Xử lý trên số (trunc + chia 10) để tránh String(số) dạng khoa học.
 */
function luffyXuLyCotCSkuSoBo10So0(
  bat: boolean,
  raw: string | number | boolean,
): string | number | boolean {
  if (!bat) return raw;
  if (raw === null || raw === undefined) return raw;
  if (typeof raw === "boolean") return raw;

  let n: number;
  if (typeof raw === "number") {
    if (!isFinite(raw)) return raw;
    n = raw;
  } else {
    let s = luffyLamSach(String(raw));
    if (s === "") return raw;
    n = Number(s);
    if (!isFinite(n)) return raw;
  }

  let x = Math.trunc(n);
  let removed = 0;
  while (removed < 10 && x !== 0 && x % 10 === 0) {
    x = x / 10;
    removed++;
  }
  return x;
}

// Expose pure helpers for Node-based regression tests without affecting Excel execution.
/**
 * Tự động tạo các sheet theo danh sách tên.
 * - Nếu sheet đã tồn tại → bỏ qua (log thông báo).
 * - Nếu chưa có → tạo mới (log thông báo).
 * Trả về số sheet đã tạo mới.
 */
function luffyTaoSheetNeuChuaCo(
  workbook: ExcelScript.Workbook,
  danhSachTen: string[],
): number {
  let soTaoMoi = 0;
  let logs: string[] = [];
  for (let ten of danhSachTen) {
    let tenTrimmed = ten.trim();
    if (tenTrimmed === "") continue;
    let sheetHienCo = workbook.getWorksheet(tenTrimmed);
    if (sheetHienCo) {
      logs.push(`Sheet "${tenTrimmed}" đã tồn tại → bỏ qua.`);
    } else {
      workbook.addWorksheet(tenTrimmed);
      soTaoMoi++;
      logs.push(`Đã tạo sheet mới: "${tenTrimmed}".`);
    }
  }
  if (logs.length > 0) {
    console.log(logs.join("\n"));
  }
  return soTaoMoi;
}

/**
 * Cài đặt cấu trúc sheet ERP-A:
 * - Ghi tiêu đề hàng 1, freeze pane, tạo table có auto-filter.
 * - Đặt công thức: A=ROW()-1, E=C&B, F=SUMIFS(ERP!N:N,ERP!O:O,E), G=D-F.
 * - Chỉ cài đặt khi hàng tiêu đề trống (tránh ghi đè dữ liệu đã có).
 */
function luffyCaiDatSheetErpA(
  workbook: ExcelScript.Workbook,
  soDongNguon: number,
  config: typeof LUFFY_CONFIG,
): void {
  let cauHinh = config.cauHinhErpA;
  let sheetA = workbook.getWorksheet(cauHinh.tenSheet);
  if (!sheetA) {
    console.log(
      `Sheet "${cauHinh.tenSheet}" không tồn tại — bỏ qua cài đặt cấu trúc cho sheet này.`,
    );
    return;
  }

  // Kiểm tra hàng tiêu đề đã có nội dung chưa (tránh ghi đè)
  let soCot = cauHinh.tieuDe.length;
  let cotCuoi = String.fromCharCode(64 + soCot); // A=65, G=71 → 7 cột = G
  let rangeTieuDeA = sheetA.getRange(
    `A${cauHinh.dongTieuDe}:${cotCuoi}${cauHinh.dongTieuDe}`,
  );

  // Định dạng Header cho sheet đối chiếu: Times New Roman, cỡ 14, không in đậm, căn giữa, màu nền
  let fmtHeaderA = rangeTieuDeA.getFormat();
  fmtHeaderA.getFont().setName(config.formatHeader.fontName);
  fmtHeaderA.getFont().setSize(config.formatHeader.fontSize);
  fmtHeaderA.getFont().setBold(config.formatHeader.bold);
  fmtHeaderA.getFont().setColor(config.formatHeader.fontColor);
  fmtHeaderA.getFill().setColor(config.formatHeader.fillColorDefault);
  fmtHeaderA.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  fmtHeaderA.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  let hangTieuDe = rangeTieuDeA.getValues()[0];
  let daCoCauTruc = hangTieuDe.some((cell) => String(cell ?? "").trim() !== "");
  if (daCoCauTruc) {
    let usedRangeA = sheetA.getUsedRange();
    if (usedRangeA) {
      usedRangeA.getFormat().getFont().setName(config.formatHeader.fontName);
      usedRangeA.getFormat().getFont().setSize(config.formatHeader.fontSize);
      usedRangeA.getFormat().autofitColumns();
    }
    console.log(
      `Sheet "${cauHinh.tenSheet}" đã có tiêu đề — chỉ cập nhật định dạng Header & Font.`,
    );
    return;
  }

  let bd = cauHinh.dongBatDauDuLieu;
  let soDong = cauHinh.soDongDuLieu > 0 ? cauHinh.soDongDuLieu : soDongNguon;
  if (soDong <= 0) soDong = 50; // fallback
  let dongCuoi = bd + soDong - 1;
  let { stt, hopDong, sku, sl, code, tong, chenhLech } = cauHinh.cot;

  // Tên sheet ERP dùng trong SUMIFS
  let tenErp = cauHinh.tenSheetErpThamChieu || config.tenSheetDich;
  // Escape tên sheet có ký tự đặc biệt
  let tenErpRef =
    tenErp.includes(" ") || tenErp.includes("-") ? `'${tenErp}'` : tenErp;
  let tenErpARef =
    cauHinh.tenSheet.includes(" ") || cauHinh.tenSheet.includes("-")
      ? `'${cauHinh.tenSheet}'`
      : cauHinh.tenSheet;
  let cotSumSl = cauHinh.cotErpSumSl;
  let cotCriteriaCode = cauHinh.cotErpCode;

  // 1. Ghi tiêu đề
  let tieuDeRow: (string | number | boolean)[][] = [
    cauHinh.tieuDe.map((t) => t),
  ];
  sheetA
    .getRange(`A${cauHinh.dongTieuDe}:${cotCuoi}${cauHinh.dongTieuDe}`)
    .setValues(tieuDeRow);

  // 2. Công thức cột A (STT = ROW()-1)
  let congThucStt: string[][] = [];
  for (let i = 0; i < soDong; i++) {
    congThucStt.push([`=ROW()-1`]);
  }
  sheetA.getRange(`${stt}${bd}:${stt}${dongCuoi}`).setFormulas(congThucStt);

  // 3. Công thức cột E (Code = C&B)
  let congThucCode: string[][] = [];
  for (let i = 0; i < soDong; i++) {
    let row = bd + i;
    congThucCode.push([`=${sku}${row}&${hopDong}${row}`]);
  }
  sheetA.getRange(`${code}${bd}:${code}${dongCuoi}`).setFormulas(congThucCode);

  // 4. Công thức cột F (Tổng = SUMIFS từ ERP)
  let congThucTong: string[][] = [];
  for (let i = 0; i < soDong; i++) {
    let row = bd + i;
    congThucTong.push([
      `=SUMIFS(${tenErpRef}!${cotSumSl}:${cotSumSl},${tenErpRef}!${cotCriteriaCode}:${cotCriteriaCode},${tenErpARef}!${code}${row})`,
    ]);
  }
  sheetA.getRange(`${tong}${bd}:${tong}${dongCuoi}`).setFormulas(congThucTong);

  // 5. Công thức cột G (Chênh lệch = D - F)
  let congThucChenhLech: string[][] = [];
  for (let i = 0; i < soDong; i++) {
    let row = bd + i;
    congThucChenhLech.push([`=${sl}${row}-${tong}${row}`]);
  }
  sheetA
    .getRange(`${chenhLech}${bd}:${chenhLech}${dongCuoi}`)
    .setFormulas(congThucChenhLech);

  // 6. Tạo Table (auto-filter) — xóa table cũ nếu có
  let tableRange = sheetA.getRange(
    `A${cauHinh.dongTieuDe}:${cotCuoi}${dongCuoi}`,
  );
  let existingTables = sheetA.getTables();
  for (let t of existingTables) {
    t.delete();
  }
  let table = sheetA.addTable(tableRange, true);
  let tenBangSanSang = cauHinh.tenSheet.replace(/[^a-zA-Z0-9_]/g, "");
  table.setName(`Table_${tenBangSanSang}`);
  table.setShowFilterButton(true);

  // 7. Freeze hàng tiêu đề
  sheetA.getFreezePanes().freezeRows(cauHinh.dongTieuDe);

  // 8. Đặt Font Times New Roman 14 và tự căn chỉnh độ rộng cột theo độ dài tiêu đề
  let fullRangeA = sheetA.getRange(`A1:${cotCuoi}${dongCuoi}`);
  fullRangeA.getFormat().getFont().setName(config.formatHeader.fontName);
  fullRangeA.getFormat().getFont().setSize(config.formatHeader.fontSize);
  fullRangeA.getFormat().autofitColumns();

  console.log(
    `Đã cài đặt cấu trúc "${cauHinh.tenSheet}": tiêu đề dòng ${cauHinh.dongTieuDe}, ` +
      `${soDong} dòng công thức (${bd}→${dongCuoi}), table + filter + freeze + autofit.`,
  );
}

const luffyTestables = {
  luffyLamSach,
  luffyHangNguonTrong,
  luffyLaySoStt,
  luffyTrongSoLuong,
  luffyXuLyCotCSkuSoBo10So0,
  luffyTachKichThuoc,
};

(
  globalThis as typeof globalThis & {
    __AUTO_LUFFY_TESTABLES__?: typeof luffyTestables;
  }
).__AUTO_LUFFY_TESTABLES__ = luffyTestables;

// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
function main(workbook: ExcelScript.Workbook) {
  const CONFIG = LUFFY_CONFIG;

  // === Bước 0: Tự động tạo sheet nếu chưa có ===
  if (CONFIG.autoTaoSheet.length > 0) {
    let soMoi = luffyTaoSheetNeuChuaCo(workbook, CONFIG.autoTaoSheet);
    if (soMoi > 0) {
      console.log(`Tổng cộng đã tạo ${soMoi} sheet mới.`);
    }
  }
  type DichVung = (typeof CONFIG.dichVung)[number];
  const laVungTuSheet1 = (
    v: DichVung,
  ): v is Extract<DichVung, { duLieuDataSheet: false }> =>
    "duLieuDataSheet" in v;
  const laVungNhapTay = (
    v: DichVung,
  ): v is Extract<DichVung, { duLieuERP: false }> => "duLieuERP" in v;
  const coCotUnique = (
    v: DichVung,
  ): v is Extract<DichVung, { unique: string }> => "unique" in v;
  const laVungCoKichThuoc = (
    v: DichVung,
  ): v is Extract<DichVung, { dai: string; rong: string; cao: string }> =>
    laVungTuSheet1(v) && "dai" in v && "rong" in v && "cao" in v;
  const cotBatDauDuLieuSheet1 = (v: DichVung): string => {
    if (laVungCoKichThuoc(v)) return v.dai;
    return v.hopDong;
  };

  let sheetNguon = workbook.getWorksheet(CONFIG.tenSheetNguon);
  if (!sheetNguon) {
    console.log(`Không tìm thấy sheet nguồn "${CONFIG.tenSheetNguon}".`);
    return;
  }

  let sheetDich = CONFIG.tenSheetDich
    ? workbook.getWorksheet(CONFIG.tenSheetDich)
    : workbook.getActiveWorksheet();
  if (!sheetDich) {
    console.log(
      CONFIG.tenSheetDich
        ? `Không tìm thấy sheet đích "${CONFIG.tenSheetDich}".`
        : "Không lấy được sheet đang mở (đích).",
    );
    return;
  }

  let headerRowIdx = luffyFindHeaderRow(
    sheetNguon,
    CONFIG.dataTable.headerKeywords,
  );
  if (headerRowIdx === -1) {
    console.log(
      `${CONFIG.tenSheetNguon}: không tìm thấy hàng tiêu đề chứa [${CONFIG.dataTable.headerKeywords.join(", ")}] trong cột A.`,
    );
    return;
  }
  let headerRowNum = headerRowIdx + 1;
  console.log(`${CONFIG.tenSheetNguon}: hàng tiêu đề dòng ${headerRowNum}.`);

  let dongNguon = headerRowNum + 1;
  let ketQuaStt = luffyTimVungTheoCotStt(
    sheetNguon,
    dongNguon,
    workbook,
    CONFIG.cotSoThuTu,
  );
  if (ketQuaStt.ok === false) {
    console.log(ketQuaStt.loi);
    return;
  }
  let dongNguonCuoi = ketQuaStt.lastRow;
  console.log(
    `${CONFIG.tenSheetNguon}: vùng theo STT (${CONFIG.cotSoThuTu}) dòng ${dongNguon} → ${dongNguonCuoi} (1…n liên tục, dừng khi STT trống).`,
  );

  let n = dongNguonCuoi - dongNguon + 1;
  let {
    hopDong: cHopDong,
    sku: cSku,
    sl: cSoLuong,
    kichThuoc: cKichThuoc,
  } = CONFIG.nguon;
  let vungTuSheet1 = CONFIG.dichVung.find((v) => laVungTuSheet1(v));
  let gopKichThuoc =
    vungTuSheet1 !== undefined &&
    "dai" in vungTuSheet1 &&
    "rong" in vungTuSheet1 &&
    "cao" in vungTuSheet1;
  let valsHCheck = sheetNguon
    .getRange(`${cSoLuong}${dongNguon}:${cSoLuong}${dongNguonCuoi}`)
    .getValues();
  let chiSoThieuSl = -1;
  for (let i = 0; i < n; i++) {
    if (luffyTrongSoLuong(valsHCheck[i][0])) {
      chiSoThieuSl = i;
      break;
    }
  }
  if (chiSoThieuSl >= 0) {
    console.log(
      `Lỗi: dòng ${dongNguon + chiSoThieuSl} (STT ${chiSoThieuSl + 1}) số lượng (${cSoLuong}) trống.`,
    );
    return;
  }

  if (CONFIG.dongCongThucUnique <= CONFIG.dongTieuDeCotE) {
    console.log(
      `dongCongThucUnique phải lớn hơn dongTieuDeCotE để không ghi đè tiêu đề (E/F dòng ${CONFIG.dongTieuDeCotE}).`,
    );
    return;
  }

  let cotHopDong = sheetNguon
    .getRange(`${cHopDong}${dongNguon}:${cHopDong}${dongNguonCuoi}`)
    .getValues();
  let cotSku = sheetNguon
    .getRange(`${cSku}${dongNguon}:${cSku}${dongNguonCuoi}`)
    .getValues();
  let cotSl = sheetNguon
    .getRange(`${cSoLuong}${dongNguon}:${cSoLuong}${dongNguonCuoi}`)
    .getValues();
  let cotKichThuoc: (string | number | boolean)[][] = [];
  if (gopKichThuoc) {
    cotKichThuoc = sheetNguon
      .getRange(`${cKichThuoc}${dongNguon}:${cKichThuoc}${dongNguonCuoi}`)
      .getValues();
  }

  let nBanDau = n;
  if (CONFIG.boQuaDongTrongNguon) {
    let dau = 0;
    while (
      dau < n &&
      luffyHangNguonTrong(cotHopDong[dau][0], cotSku[dau][0], cotSl[dau][0])
    ) {
      dau++;
    }
    let cuoi = n - 1;
    while (
      cuoi >= dau &&
      luffyHangNguonTrong(cotHopDong[cuoi][0], cotSku[cuoi][0], cotSl[cuoi][0])
    ) {
      cuoi--;
    }
    if (cuoi < dau) {
      console.log(
        `${CONFIG.tenSheetNguon}: sau khi bỏ dòng trống đầu/cuối không còn dòng dữ liệu (B, C, H).`,
      );
      return;
    }
    if (dau > 0 || cuoi < n - 1) {
      console.log(
        `Đã bỏ ${dau} dòng trống đầu vùng + ${n - 1 - cuoi} dòng trống cuối (trên ${nBanDau} dòng vùng ${CONFIG.tenSheetNguon}).`,
      );
    }
    cotHopDong = cotHopDong.slice(dau, cuoi + 1);
    cotSku = cotSku.slice(dau, cuoi + 1);
    cotSl = cotSl.slice(dau, cuoi + 1);
    if (gopKichThuoc) {
      cotKichThuoc = cotKichThuoc.slice(dau, cuoi + 1);
    }
    n = cuoi - dau + 1;
  }

  let bd = CONFIG.dongBatDau;
  let dongDichCuoiNguon = bd + n - 1;
  let vungNhapTay = CONFIG.dichVung.find(laVungNhapTay);
  let dongDichQuetCuoi = luffyLayDongCuoiQuet(sheetDich, bd, workbook);
  let dongNhapTayCuoi = vungNhapTay
    ? luffyLayDongCuoiCoDuLieuTrongVung(
        sheetDich,
        bd,
        dongDichQuetCuoi,
        vungNhapTay.hopDong,
        vungNhapTay.sl,
      )
    : null;
  let dongDichCuoiNhapTay = dongNhapTayCuoi ?? dongDichCuoiNguon;

  let giaTriNguon: (string | number | boolean)[][] = [];
  for (let r = 0; r < n; r++) {
    let hd = cotHopDong[r][0];
    let sku = cotSku[r][0];
    let sl = cotSl[r][0];
    if (gopKichThuoc) {
      let [dai, rong, cao] = luffyTachKichThuoc(cotKichThuoc[r][0]);
      giaTriNguon.push([dai, rong, cao, hd, sku, sl]);
    } else {
      giaTriNguon.push([hd, sku, sl]);
    }
  }

  let dongXoaUnique = CONFIG.dongCongThucUnique + 1;
  let cSttCodeUnique = CONFIG.cotSttCodeUnique;
  let cH = CONFIG.cotSoLuongConLai;
  let {
    cotKetQua: cF,
    cotSum: cJ,
    cotKey: cK,
    cotTieuChi: cE,
  } = CONFIG.sumIfsTheoMaDoc;
  sheetDich
    .getRange(`${cF}${bd}:${cF}${CONFIG.xoaCotEdenDong}`)
    .clear(ExcelScript.ClearApplyTo.contents);
  sheetDich
    .getRange(`${cH}${bd}:${cH}${CONFIG.xoaCotEdenDong}`)
    .clear(ExcelScript.ClearApplyTo.contents);
  sheetDich
    .getRange(
      `${cSttCodeUnique}${dongXoaUnique}:${cSttCodeUnique}${CONFIG.xoaCotEdenDong}`,
    )
    .clear(ExcelScript.ClearApplyTo.contents);

  for (let v of CONFIG.dichVung) {
    if (laVungTuSheet1(v)) {
      let cotPhaiXoa =
        gopKichThuoc && laVungCoKichThuoc(v)
          ? CONFIG.cotTamKichThuoc.cao
          : v.code;
      sheetDich
        .getRange(
          `${cotBatDauDuLieuSheet1(v)}${bd}:${cotPhaiXoa}${CONFIG.xoaCotEdenDong}`,
        )
        .clear(ExcelScript.ClearApplyTo.contents);
    } else {
      sheetDich
        .getRange(`${v.code}${bd}:${v.code}${CONFIG.xoaCotEdenDong}`)
        .clear(ExcelScript.ClearApplyTo.contents);
    }
    if (coCotUnique(v)) {
      sheetDich
        .getRange(
          `${v.unique}${dongXoaUnique}:${v.unique}${CONFIG.xoaCotEdenDong}`,
        )
        .clear(ExcelScript.ClearApplyTo.contents);
    }
  }

  for (let v of CONFIG.dichVung) {
    let dongVungCuoi = laVungTuSheet1(v)
      ? dongDichCuoiNguon
      : dongDichCuoiNhapTay;
    if (laVungTuSheet1(v)) {
      if (gopKichThuoc && laVungCoKichThuoc(v)) {
        let { dai: cP, rong: cQ, cao: cR } = CONFIG.cotTamKichThuoc;
        let tamKT: (string | number | boolean)[][] = [];
        let lmn: (string | number | boolean)[][] = [];
        for (let r = 0; r < n; r++) {
          let row = giaTriNguon[r];
          tamKT.push([row[0], row[1], row[2]]);
          lmn.push([row[3], row[4], row[5]]);
        }
        sheetDich.getRange(`${cP}${bd}:${cR}${dongVungCuoi}`).setValues(tamKT);
        sheetDich
          .getRange(`${v.hopDong}${bd}:${v.sl}${dongVungCuoi}`)
          .setValues(lmn);
      } else {
        sheetDich
          .getRange(`${cotBatDauDuLieuSheet1(v)}${bd}:${v.sl}${dongVungCuoi}`)
          .setValues(giaTriNguon);
      }
    }
    let congThucCode: string[][] = [];
    for (let i = 0; i <= dongVungCuoi - bd; i++) {
      let excelRow = bd + i;
      congThucCode.push([`=${v.sku}${excelRow}&${v.hopDong}${excelRow}`]);
    }
    sheetDich
      .getRange(`${v.code}${bd}:${v.code}${dongVungCuoi}`)
      .setFormulas(congThucCode);
    if (coCotUnique(v)) {
      let vungCode = `${v.code}${bd}:${v.code}${dongVungCuoi}`;
      let congThucE = CONFIG.uniqueEFilterBoTrong
        ? `=UNIQUE(FILTER(${vungCode},${vungCode}<>""))`
        : `=UNIQUE(${vungCode})`;
      sheetDich
        .getRange(`${v.unique}${CONFIG.dongCongThucUnique}`)
        .setFormula(congThucE);
    }
  }

  // === Đặt công thức FILTER tại A3 trên sheet ERP lấy dữ liệu từ sheet đối chiếu (ERP-A) ===
  let cfgFilterA3 = CONFIG.congThucFilterA3ERP;
  let tenErpA = CONFIG.cauHinhErpA.tenSheet;
  let tenErpARef =
    tenErpA.includes(" ") || tenErpA.includes("-") ? `'${tenErpA}'` : tenErpA;
  let congThucA3 = `=FILTER(${tenErpARef}!${cfgFilterA3.vungLay},${tenErpARef}!${cfgFilterA3.vungKiemTra}<>"","")`;
  sheetDich.getRange(cfgFilterA3.oBatDau).setFormula(congThucA3);

  if (
    gopKichThuoc &&
    vungTuSheet1 !== undefined &&
    laVungCoKichThuoc(vungTuSheet1)
  ) {
    let v = vungTuSheet1;
    let { dai: cP, rong: cQ, cao: cR } = CONFIG.cotTamKichThuoc;
    let cO = v.code;
    let oDai = v.dai;
    let oRong = v.rong;
    let oCao = v.cao;
    let dongIJKCuoi = Math.max(dongDichCuoiNhapTay, dongDichCuoiNguon);
    let endNguon = dongDichCuoiNguon;
    let congThucI: string[][] = [];
    let congThucJ: string[][] = [];
    let congThucK: string[][] = [];
    for (let excelRow = bd; excelRow <= dongIJKCuoi; excelRow++) {
      congThucI.push([
        `=IF($${cE}${excelRow}="",$${cP}${excelRow},XLOOKUP($${cE}${excelRow},$${cO}$${bd}:$${cO}$${endNguon},$${cP}$${bd}:$${cP}$${endNguon},"",0,-1))`,
      ]);
      congThucJ.push([
        `=IF($${cE}${excelRow}="",$${cQ}${excelRow},XLOOKUP($${cE}${excelRow},$${cO}$${bd}:$${cO}$${endNguon},$${cQ}$${bd}:$${cQ}$${endNguon},"",0,-1))`,
      ]);
      congThucK.push([
        `=IF($${cE}${excelRow}="",$${cR}${excelRow},XLOOKUP($${cE}${excelRow},$${cO}$${bd}:$${cO}$${endNguon},$${cR}$${bd}:$${cR}$${endNguon},"",0,-1))`,
      ]);
    }
    sheetDich
      .getRange(`${oDai}${bd}:${oDai}${dongIJKCuoi}`)
      .setFormulas(congThucI);
    sheetDich
      .getRange(`${oRong}${bd}:${oRong}${dongIJKCuoi}`)
      .setFormulas(congThucJ);
    sheetDich
      .getRange(`${oCao}${bd}:${oCao}${dongIJKCuoi}`)
      .setFormulas(congThucK);
  }

  if (vungNhapTay) {
    let oUniqueDau = `${vungNhapTay.unique}${CONFIG.dongCongThucUnique}`;
    sheetDich
      .getRange(`${cSttCodeUnique}${CONFIG.dongCongThucUnique}`)
      .setFormula(`=IF(${oUniqueDau}#="","",SEQUENCE(ROWS(${oUniqueDau}#)))`);
  }

  let dongKetQuaCuoi = dongDichCuoiNhapTay;
  let congThucF: string[][] = [];
  let congThucH: string[][] = [];
  for (let i = 0; i <= dongKetQuaCuoi - bd; i++) {
    let excelRow = bd + i;
    congThucF.push([
      `=SUMIFS($${cJ}$${bd}:$${cJ}$${dongDichCuoiNguon},$${cK}$${bd}:$${cK}$${dongDichCuoiNguon},${cE}${excelRow})`,
    ]);
    if (vungNhapTay) {
      congThucH.push([`=${vungNhapTay.sl}${excelRow}-${cF}${excelRow}`]);
    }
  }
  sheetDich
    .getRange(`${cF}${bd}:${cF}${dongKetQuaCuoi}`)
    .setFormulas(congThucF);
  if (vungNhapTay) {
    sheetDich
      .getRange(`${cH}${bd}:${cH}${dongKetQuaCuoi}`)
      .setFormulas(congThucH);
  }
  let ketQuaRange = sheetDich.getRange(`${cF}${bd}:${cF}${dongKetQuaCuoi}`);
  ketQuaRange.clearAllConditionalFormats();
  let cfHBangKhong = ketQuaRange.addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom,
  );
  cfHBangKhong.getCustom().getRule().setFormula(`=$${cH}${bd}=0`);
  cfHBangKhong.getCustom().getFormat().getFill().setColor("#00B050");
  cfHBangKhong.getCustom().getFormat().getFont().setColor("#FFFFFF");
  cfHBangKhong.setPriority(1);
  cfHBangKhong.setStopIfTrue(true);
  let cfGBangKhong = ketQuaRange.addConditionalFormat(
    ExcelScript.ConditionalFormatType.cellValue,
  );
  cfGBangKhong.getCellValue().setRule({
    operator: ExcelScript.ConditionalCellValueOperator.equalTo,
    formula1: "0",
  });
  cfGBangKhong.getCellValue().getFormat().getFill().setColor("#FF0000");
  cfGBangKhong.getCellValue().getFormat().getFont().setColor("#FFFFFF");
  cfGBangKhong.setPriority(2);

  // === Cảnh báo tô màu dữ liệu trùng lặp ở cột O (Code) ===
  let cfgO = CONFIG.canhBaoTrungLapCotO;
  if (cfgO && cfgO.bat) {
    let cotO = cfgO.cot;
    let maxDongO = Math.max(dongDichCuoiNguon, dongDichCuoiNhapTay, bd);
    let rangeCotO = sheetDich.getRange(`${cotO}${bd}:${cotO}${maxDongO}`);
    rangeCotO.clearAllConditionalFormats();

    let cfTrungLapO = rangeCotO.addConditionalFormat(
      ExcelScript.ConditionalFormatType.custom,
    );
    cfTrungLapO
      .getCustom()
      .getRule()
      .setFormula(
        `=AND($${cotO}${bd}<>"",COUNTIF($${cotO}$${bd}:$${cotO}$${maxDongO},$${cotO}${bd})>1)`,
      );
    cfTrungLapO
      .getCustom()
      .getFormat()
      .getFill()
      .setColor(cfgO.mauNenCanhBao);
    cfTrungLapO
      .getCustom()
      .getFormat()
      .getFont()
      .setColor(cfgO.mauChuCanhBao);
  }

  if (CONFIG.chuyenCotCSkuThanhSoBo10So0Cuoi) {
    if (vungNhapTay) {
      let cotNhapTay = vungNhapTay.sl;
      let rngC = sheetDich.getRange(
        `${cotNhapTay}${bd}:${cotNhapTay}${dongDichCuoiNhapTay}`,
      );
      let valsC = rngC.getValues();
      let outC: (string | number | boolean)[][] = [];
      for (let r = 0; r < valsC.length; r++) {
        outC.push([luffyXuLyCotCSkuSoBo10So0(true, valsC[r][0])]);
      }
      rngC.setValues(outC);
    }
  }

  let cotTomTatNguon = gopKichThuoc
    ? `${cHopDong},${cSku},${cSoLuong},${cKichThuoc} (kích thước → dài/rộng/cao)`
    : `${cHopDong},${cSku},${cSoLuong}`;
  let tomTatVung = CONFIG.dichVung
    .map((v) => {
      let dongVungCuoi = laVungTuSheet1(v)
        ? dongDichCuoiNguon
        : dongDichCuoiNhapTay;
      return laVungTuSheet1(v)
        ? `${cotBatDauDuLieuSheet1(v)}${bd}:${v.sl}${dongVungCuoi} (Sheet1) + ${v.code}=${v.sku}&${v.hopDong}`
        : `${v.code}=${v.sku}&${v.hopDong} (A:C tay) + ${v.unique}=${CONFIG.uniqueEFilterBoTrong ? "UNIQUE(FILTER(D))" : "UNIQUE(D)"} + ${CONFIG.cotSttCodeUnique}=STT(${v.unique}#)`;
    })
    .join("; ");
  // === Ghi tiêu đề, công thức SUBTOTAL G1/N1, định dạng Header, Freeze Panes và AutoFilter cho sheet ERP ===
  if (CONFIG.tieuDeERP && CONFIG.tieuDeERP.length > 0) {
    let tieuDeERP = CONFIG.tieuDeERP;
    let cotCuoiERP = String.fromCharCode(64 + tieuDeERP.length);
    let dongTieuDeERP = CONFIG.dongTieuDeCotE; // Dòng 2
    let rangeTieuDeERP = sheetDich.getRange(
      `A${dongTieuDeERP}:${cotCuoiERP}${dongTieuDeERP}`,
    );
    rangeTieuDeERP.setValues([tieuDeERP]);

    // 1. Định dạng font, cỡ chữ, căn giữa cho hàng tiêu đề
    let fmtERP = rangeTieuDeERP.getFormat();
    fmtERP.getFont().setName(CONFIG.formatHeader.fontName);
    fmtERP.getFont().setSize(CONFIG.formatHeader.fontSize);
    fmtERP.getFont().setBold(CONFIG.formatHeader.bold);
    fmtERP.getFont().setColor(CONFIG.formatHeader.fontColor);
    fmtERP.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    fmtERP.setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    // 2. Màu nền ô phân biệt riêng biệt theo nhóm cột: A-F, G-H, I-K, L-R
    let fills = CONFIG.formatHeader.fillColorsERP;
    sheetDich
      .getRange(`A${dongTieuDeERP}:F${dongTieuDeERP}`)
      .getFormat()
      .getFill()
      .setColor(fills.nhomAF);
    sheetDich
      .getRange(`G${dongTieuDeERP}:H${dongTieuDeERP}`)
      .getFormat()
      .getFill()
      .setColor(fills.nhomGH);
    sheetDich
      .getRange(`I${dongTieuDeERP}:K${dongTieuDeERP}`)
      .getFormat()
      .getFill()
      .setColor(fills.nhomIK);
    sheetDich
      .getRange(`L${dongTieuDeERP}:${cotCuoiERP}${dongTieuDeERP}`)
      .getFormat()
      .getFill()
      .setColor(fills.nhomLR);

    // 3. Đặt công thức SUBTOTAL cho G1 và N1 (=SUBTOTAL(9,G3:G1000) & =SUBTOTAL(9,N3:N1000))
    let cellG1 = sheetDich.getRange("G1");
    let cellN1 = sheetDich.getRange("N1");
    cellG1.setFormula("=SUBTOTAL(9,G3:G1000)");
    cellN1.setFormula("=SUBTOTAL(9,N3:N1000)");

    // Định dạng ô SUBTOTAL G1 & N1 (Font Times New Roman, cỡ 14, in đậm, căn phải, màu nền nổi bật)
    [cellG1, cellN1].forEach((cell) => {
      let fmtCell = cell.getFormat();
      fmtCell.getFont().setName(CONFIG.formatHeader.fontName);
      fmtCell.getFont().setSize(CONFIG.formatHeader.fontSize);
      fmtCell.getFont().setBold(true);
      fmtCell.getFont().setColor(CONFIG.formatHeader.subtotalFontColor);
      fmtCell.getFill().setColor(CONFIG.formatHeader.subtotalFillColor);
      fmtCell.setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
      fmtCell.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    });

    // 4. Freeze tiêu đề hàng 2 trên sheet ERP
    sheetDich.getFreezePanes().freezeRows(dongTieuDeERP);

    // 5. Bật bộ lọc dữ liệu (AutoFilter) trên sheet ERP
    let dongCuoiERP = Math.max(dongDichCuoiNhapTay, dongDichCuoiNguon);
    if (dongCuoiERP >= dongTieuDeERP) {
      let rangeFilterERP = sheetDich.getRange(
        `A${dongTieuDeERP}:${cotCuoiERP}${dongCuoiERP}`,
      );
      sheetDich.getAutoFilter().apply(rangeFilterERP);
    }

    // 6. Áp dụng Font Times New Roman 14 và tự căn chỉnh độ rộng cột theo độ dài tiêu đề
    let maxDongERP = Math.max(dongDichCuoiNhapTay, dongDichCuoiNguon, 3);
    let fullRangeERP = sheetDich.getRange(`A1:${cotCuoiERP}${maxDongERP}`);
    fullRangeERP.getFormat().getFont().setName(CONFIG.formatHeader.fontName);
    fullRangeERP.getFormat().getFont().setSize(CONFIG.formatHeader.fontSize);
    fullRangeERP.getFormat().autofitColumns();
  }

  // === Cài đặt cấu trúc sheet đối chiếu (${CONFIG.cauHinhErpA.tenSheet}) ===
  luffyCaiDatSheetErpA(workbook, n, CONFIG);

  console.log(
    `Luffy: ${n} dòng ${CONFIG.tenSheetNguon} (${cotTomTatNguon}) → "${sheetDich.getName()}". ${cF}=SUMIFS(${cJ},${cK},${cE}), ${cH}=${vungNhapTay ? vungNhapTay.sl : "C"}-${cF}. ${tomTatVung}. Giữ thứ tự ${CONFIG.tenSheetNguon}.`,
  );
}
