/**
 * ===================================================================
 * Auto_Nami_Support – Bổ sung cột thông số cho sheet kết quả Auto_Nami
 * ===================================================================
 *
 * Cùng layout cột với Auto_Nami (A–T). Script CHỈ ghi các cột thông số
 * M–S khi ô còn trống — không thay A–L/T và không ghi đè dữ liệu Auto_Nami
 * đã điền (tra cứu sheet thông số khi chạy với thongSo.bat=true).
 *
 * Cột nguồn tra cứu: F (SKU), G (Tên hàng), J (Số lượng nhập — cho công thức).
 * Cột bổ sung: M–S (Trọng lượng … Tổng thể tích). Cột T (Vị trí) do Auto_Nami ghi.
 *
 * Ngoài ra, khi có dòng SKU được thêm vào sau khi Auto_Nami đã chạy, script
 * cập nhật lại dòng cuối trong công thức SUBTOTAL hàng 1 (J1/N1/S1) và dựng
 * lại hai bảng thống kê (từ cột V, tiêu đề bắt đầu ở hàng 3 — dưới vùng freeze
 * panes 2 dòng đầu; cột U trống tạo khoảng cách với cột Vị trí T) theo dữ liệu
 * hiện tại — bao trùm cả dòng mới lẫn mã phân loại / thị trường mới chưa có
 * trong bảng cũ.
 */

interface NamiThongSoItem {
  trongLuong: number;
  dai: number;
  rong: number;
  cao: number;
  theTich: number;
}

type NamiSupportThongSoTheoTenEntry = {
  tenChuan: string;
  tokens: string[];
  thongSo: NamiThongSoItem;
  coDuLieu: boolean;
};

type NamiSupportTraCuuTenMoConfig = {
  bat: boolean;
  diemToiThieu: number;
  tyLeTokenKhopToiThieu: number;
  soTokenKhopToiDa: number;
};

type NamiThongSoTraCuuResult = {
  thongSo?: NamiThongSoItem;
  nguon: "ma" | "ten-chinh-xac" | "ten-mo" | "khong-co";
};

const NAMI_SUPPORT_CONFIG = {
  tenSheetThongSo: "2026",
  dongBatDauDuLieu: 3, // Hàng 1 subtotal, hàng 2 header (Auto_Nami)

  traCuuTenMo: {
    bat: true,
    diemToiThieu: 0.55,
    tyLeTokenKhopToiThieu: 0.5,
    soTokenKhopToiDa: 3,
  } as NamiSupportTraCuuTenMoConfig,

  /** Cột Auto_Nami đã xuất — chỉ đọc, không ghi */
  cotNguon: {
    thiTruong: "D",
    maPhanLoai: "E",
    sku: "F",
    tenHang: "G",
    soLuongNhap: "J",
  },

  /** Cột thông số M–S (khớp Auto_Nami) — chỉ ghi khi ô trống */
  cotChiSo: {
    trongLuong: "M",
    tongTrongLuong: "N",
    dai: "O",
    rong: "P",
    cao: "Q",
    theTich: "R",
    tongTheTich: "S",
  },

  /** Cập nhật dòng cuối trong công thức tổng hợp của Auto_Nami khi có SKU thêm */
  congThucTongHop: {
    bat: true,
    /** Ô SUBTOTAL hàng 1 (khớp Auto_Nami: J1, N1, S1) */
    oSubtotal: ["J1", "N1", "S1"],
  },

  /**
   * Dựng lại hai bảng thống kê của Auto_Nami theo dữ liệu hiện tại — bao trùm
   * cả dòng SKU thêm và mã phân loại / thị trường mới chưa có trong bảng cũ.
   */
  thongKeKetQua: {
    bat: true,
    cotBatDau: "V", // Khớp thongKeKetQua.cotBatDau của Auto_Nami (bảng V–AA; U trống).
    dongBatDauBang: 3, // Khớp thongKeKetQua.dongBatDau của Auto_Nami — tiêu đề dưới 2 dòng freeze panes.
    soCot: 6,
    heSoChiaTheTich: 68,
    thuTuMaPhanLoai: ["F", "A", "S", "C", "T", "M"],
  },

  format: {
    fontName: "Times New Roman",
    fontSize: 14,
    subtotalFillColor: "#FAB0B5",
    headerFillColor: "#4F81BD",
  },
};

/** Cập nhật dòng cuối trong công thức =SUBTOTAL(9, J3:J68) → =SUBTOTAL(9, J3:J<dongCuoi>) */
const namiSupportCapNhatCongThucSubtotal = (
  congThuc: string,
  dongCuoi: number,
): string | null => {
  const match = String(congThuc ?? "").match(
    /^=\s*SUBTOTAL\(\s*(\d+)\s*[,;]\s*([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)\s*\)\s*$/i,
  );
  if (!match) return null;
  const [, maHam, cotDau, dongDau, cotCuoi, dongCuoiCu] = match;
  if (Number(dongCuoiCu) === dongCuoi) return null;
  return `=SUBTOTAL(${maHam}, ${cotDau}${dongDau}:${cotCuoi}${dongCuoi})`;
};

const namiSupportDoiCotSangSo = (cot: string): number => {
  let result = 0;
  const normalized = cot.trim().toUpperCase();
  for (let i = 0; i < normalized.length; i++) {
    result = result * 26 + (normalized.charCodeAt(i) - 64);
  }
  return result;
};

const namiSupportDoiSoSangCot = (soCot: number): string => {
  let result = "";
  let n = soCot;
  while (n > 0) {
    const phanDu = (n - 1) % 26;
    result = String.fromCharCode(65 + phanDu) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
};

type NamiSupportCapThiTruongMa = {
  thiTruong: string;
  maPhanLoai: string;
};

type NamiSupportBangThongKe = {
  thongKeMaStartRow: number;
  thongKeMaRows: (string | number | boolean)[][];
  soMaPhanLoai: number;
  thongKeThiTruongStartRow: number;
  thongKeThiTruongRows: (string | number | boolean)[][];
  dsThiTruongMaPhanLoai: NamiSupportCapThiTruongMa[];
};

const namiSupportTaoChuoiCongThucExcel = (value: string): string =>
  `"${value.replace(/"/g, '""')}"`;

/** Sắp xếp mã phân loại theo thứ tự cấu hình (F, A, S, C, T, M), KHAC cuối cùng — khớp Auto_Nami */
const namiSupportSapXepMaPhanLoai = (
  a: string,
  b: string,
  thuTuCauHinh: string[],
): number => {
  const thuTu = [...thuTuCauHinh, "KHAC"];
  const idxA = thuTu.indexOf(a);
  const idxB = thuTu.indexOf(b);
  if (idxA >= 0 || idxB >= 0) {
    if (idxA < 0) return 1;
    if (idxB < 0) return -1;
    if (idxA !== idxB) return idxA - idxB;
  }
  return a.localeCompare(b);
};

/**
 * Dựng dữ liệu hai bảng thống kê (giống ghiThongKeKetQua của Auto_Nami) từ danh
 * sách thị trường / mã phân loại đọc ở cột D–E — gồm cả mã phân loại mới.
 */
const namiSupportTaoBangThongKe = (
  duLieu: NamiSupportCapThiTruongMa[],
  dongCuoi: number,
  dongBatDau: number,
  cotBatDau: string,
  heSoChiaTheTich: number,
  thuTuMaPhanLoai: string[],
  dongBatDauBang: number = 3,
): NamiSupportBangThongKe | null => {
  const tapMaPhanLoai = new Set<string>();
  const tapThiTruongMaPhanLoai = new Set<string>();
  for (const item of duLieu) {
    const thiTruong = String(item.thiTruong ?? "").trim();
    const maPhanLoai = String(item.maPhanLoai ?? "").trim();
    if (!maPhanLoai) continue;
    tapMaPhanLoai.add(maPhanLoai);
    if (thiTruong) {
      tapThiTruongMaPhanLoai.add(`${thiTruong}\u0001${maPhanLoai}`);
    }
  }

  const dsMaPhanLoai = Array.from(tapMaPhanLoai).sort((a, b) =>
    namiSupportSapXepMaPhanLoai(a, b, thuTuMaPhanLoai),
  );
  const dsThiTruongMaPhanLoai = Array.from(tapThiTruongMaPhanLoai)
    .map((key) => {
      const parts = key.split("\u0001");
      return { thiTruong: parts[0], maPhanLoai: parts[1] };
    })
    .sort(
      (a, b) =>
        a.thiTruong.localeCompare(b.thiTruong) ||
        namiSupportSapXepMaPhanLoai(a.maPhanLoai, b.maPhanLoai, thuTuMaPhanLoai),
    );
  if (dsMaPhanLoai.length === 0 && dsThiTruongMaPhanLoai.length === 0) {
    return null;
  }

  const heSoChia = heSoChiaTheTich > 0 ? heSoChiaTheTich : 68;
  const vung = (cot: string): string =>
    `$${cot}$${dongBatDau}:$${cot}$${dongCuoi}`;
  // Tiêu đề bảng thống kê nằm dưới vùng freeze panes (2 dòng đầu).
  const thongKeMaStartRow = Math.max(dongBatDauBang, 1);

  const thongKeMaRows: (string | number | boolean)[][] = [
    ["THỐNG KÊ THEO MÃ PHÂN LOẠI", "", "", ""],
    ["Mã phân loại", "Số lượng nhập", "Tổng trọng lượng", "Tổng thể tích"],
  ];
  dsMaPhanLoai.forEach((maPhanLoai) => {
    thongKeMaRows.push([
      maPhanLoai,
      `=ROUND(SUMIF(${vung("E")},${namiSupportTaoChuoiCongThucExcel(
        maPhanLoai,
      )},${vung("J")}),0)`,
      `=SUMIF(${vung("E")},${namiSupportTaoChuoiCongThucExcel(
        maPhanLoai,
      )},${vung("N")})`,
      `=SUMIF(${vung("E")},${namiSupportTaoChuoiCongThucExcel(
        maPhanLoai,
      )},${vung("S")})`,
    ]);
  });

  const thongKeThiTruongStartRow =
    thongKeMaStartRow + thongKeMaRows.length + 1;
  const cot4 = namiSupportDoiSoSangCot(namiSupportDoiCotSangSo(cotBatDau) + 4);
  const thongKeThiTruongRows: (string | number | boolean)[][] = [
    ["THỐNG KÊ THỊ TRƯỜNG THEO MÃ PHÂN LOẠI", "", "", "", "", ""],
    [
      "Thị trường",
      "Mã phân loại",
      "Số lượng nhập",
      "Tổng trọng lượng",
      "Tổng thể tích",
      `Tổng thể tích / ${heSoChia}`,
    ],
  ];
  dsThiTruongMaPhanLoai.forEach((item, idx) => {
    const rowIndex = thongKeThiTruongStartRow + 2 + idx;
    thongKeThiTruongRows.push([
      item.thiTruong,
      item.maPhanLoai,
      `=ROUND(SUMIFS(${vung("J")},${vung("D")},${namiSupportTaoChuoiCongThucExcel(
        item.thiTruong,
      )},${vung("E")},${namiSupportTaoChuoiCongThucExcel(item.maPhanLoai)}),0)`,
      `=SUMIFS(${vung("N")},${vung("D")},${namiSupportTaoChuoiCongThucExcel(
        item.thiTruong,
      )},${vung("E")},${namiSupportTaoChuoiCongThucExcel(item.maPhanLoai)})`,
      `=SUMIFS(${vung("S")},${vung("D")},${namiSupportTaoChuoiCongThucExcel(
        item.thiTruong,
      )},${vung("E")},${namiSupportTaoChuoiCongThucExcel(item.maPhanLoai)})`,
      `=${cot4}${rowIndex}/${heSoChia}`,
    ]);
  });

  return {
    thongKeMaStartRow,
    thongKeMaRows,
    soMaPhanLoai: dsMaPhanLoai.length,
    thongKeThiTruongStartRow,
    thongKeThiTruongRows,
    dsThiTruongMaPhanLoai,
  };
};

const namiSupportTestables = {
  namiSupportCapNhatCongThucSubtotal,
  namiSupportSapXepMaPhanLoai,
  namiSupportTaoBangThongKe,
};

(
  globalThis as typeof globalThis & {
    __AUTO_NAMI_SUPPORT_TESTABLES__?: typeof namiSupportTestables;
  }
).__AUTO_NAMI_SUPPORT_TESTABLES__ = namiSupportTestables;

function main(workbook: ExcelScript.Workbook) {
  const CONFIG = NAMI_SUPPORT_CONFIG;
  const sheetDich = workbook.getActiveWorksheet();

  const usedRange = sheetDich.getUsedRange();
  if (!usedRange) {
    console.log("Sheet hiện tại không có dữ liệu.");
    return;
  }

  const timDongCuoiTheoCot = (cot: string): number =>
    sheetDich
      .getRange(`${cot}1048576`)
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;

  const { cotNguon, cotChiSo } = CONFIG;
  const lastRow = Math.max(
    timDongCuoiTheoCot(cotNguon.sku),
    timDongCuoiTheoCot(cotNguon.tenHang),
    timDongCuoiTheoCot(cotNguon.soLuongNhap),
    timDongCuoiTheoCot(cotChiSo.trongLuong),
    timDongCuoiTheoCot(cotChiSo.tongTheTich),
  );
  if (lastRow < CONFIG.dongBatDauDuLieu) {
    console.log("Không có dòng dữ liệu để bổ sung thông số.");
    return;
  }

  const lamSach = (giaTri: string | number | boolean): string =>
    String(giaTri)
      .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
      .trim();

  const oTrong = (giaTri: string | number | boolean): boolean => {
    if (giaTri === "" || giaTri === null || giaTri === undefined) return true;
    if (typeof giaTri === "string" && !lamSach(giaTri)) return true;
    return false;
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
    } else if (lastComma >= 0) {
      s = s.replace(",", ".");
    }
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
    config: NamiSupportTraCuuTenMoConfig,
  ): number =>
    Math.max(
      config.soTokenKhopToiDa,
      Math.ceil(soTokenTraCuu * config.tyLeTokenKhopToiThieu),
    );

  const thongSoTenCoDuLieu = (thongSo: NamiThongSoItem): boolean =>
    !!(
      thongSo.trongLuong ||
      thongSo.dai ||
      thongSo.rong ||
      thongSo.cao ||
      thongSo.theTich
    );

  const timThongSoTheoTenMo = (
    tenChuan: string,
    danhSach: NamiSupportThongSoTheoTenEntry[],
    config: NamiSupportTraCuuTenMoConfig,
  ): NamiThongSoItem | undefined => {
    const queryTokens = tachTokenTenSanPham(tenChuan);
    if (queryTokens.length === 0 || danhSach.length === 0) return undefined;

    const nguongGiao = layNguongTokenKhop(queryTokens.length, config);

    let best: NamiThongSoItem | undefined;
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

  const lookupThongSoTheoMa = new Map<string, NamiThongSoItem>();
  const lookupThongSoTheoTen = new Map<string, NamiThongSoItem>();
  const danhSachThongSoTheoTen: NamiSupportThongSoTheoTenEntry[] = [];
  const tapTenDaThemVaoDanhSachMo = new Set<string>();
  const sheetThongSo = workbook.getWorksheet(CONFIG.tenSheetThongSo);

  if (!sheetThongSo) {
    console.log(
      `Không tìm thấy sheet thông số "${CONFIG.tenSheetThongSo}". Không thể bổ sung cột M–S.`,
    );
    return;
  }

  const lastRowThongSo =
    sheetThongSo
      .getRange("B1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;

  if (lastRowThongSo >= 2) {
    const dataThongSo = sheetThongSo
      .getRange(`B2:H${lastRowThongSo}`)
      .getValues() as (string | number | boolean)[][];

    for (const row of dataThongSo) {
      const maHang = lamSach(row[0]);
      const tenHang = chuanHoaTenSanPham(row[1]);
      if (!maHang && !tenHang) continue;

      const trongLuong = parseSoTuExcel(row[3]);
      const dai = parseSoTuExcel(row[4]);
      const rong = parseSoTuExcel(row[5]);
      const caoGoc = parseSoTuExcel(row[6]);
      const cao = caoGoc > 0 ? caoGoc : dai > 0 && rong > 0 ? 1.1 : 0;
      const theTich = (dai * rong * cao) / 1000000;

      const thongSo: NamiThongSoItem = {
        trongLuong: trongLuong ? Number(trongLuong.toFixed(4)) : 0,
        dai: dai ? Number(dai.toFixed(4)) : 0,
        rong: rong ? Number(rong.toFixed(4)) : 0,
        cao: cao ? Number(cao.toFixed(4)) : 0,
        theTich: theTich ? Number(theTich.toFixed(3)) : 0,
      };

      if (maHang) lookupThongSoTheoMa.set(maHang, thongSo);

      if (tenHang) {
        const thongSoMoiCoDuLieu = thongSoTenCoDuLieu(thongSo);
        const thongSoCu = lookupThongSoTheoTen.get(tenHang);
        const thongSoCuRong = !thongSoCu || !thongSoTenCoDuLieu(thongSoCu);
        if (thongSoCuRong || thongSoMoiCoDuLieu) {
          lookupThongSoTheoTen.set(tenHang, thongSo);
          const entryMoi = {
            tenChuan: tenHang,
            tokens: tachTokenTenSanPham(tenHang),
            thongSo,
            coDuLieu: thongSoMoiCoDuLieu,
          };
          if (tapTenDaThemVaoDanhSachMo.has(tenHang)) {
            const idx = danhSachThongSoTheoTen.findIndex(
              (item) => item.tenChuan === tenHang,
            );
            if (idx >= 0) danhSachThongSoTheoTen[idx] = entryMoi;
          } else {
            tapTenDaThemVaoDanhSachMo.add(tenHang);
            danhSachThongSoTheoTen.push(entryMoi);
          }
        }
      }
    }
  }

  const layThongSo = (
    sku: string | number | boolean,
    tenHang: string | number | boolean,
  ): NamiThongSoTraCuuResult => {
    const ma = lamSach(sku);
    if (ma) {
      const theoMa = lookupThongSoTheoMa.get(ma);
      if (theoMa) return { thongSo: theoMa, nguon: "ma" };
    }

    const tenChuan = chuanHoaTenSanPham(tenHang);
    if (!tenChuan) return { nguon: "khong-co" };

    const theoTenChinhXac = lookupThongSoTheoTen.get(tenChuan);
    if (theoTenChinhXac) {
      return { thongSo: theoTenChinhXac, nguon: "ten-chinh-xac" };
    }

    if (!CONFIG.traCuuTenMo.bat) {
      return { nguon: "khong-co" };
    }

    const theoTenMo = timThongSoTheoTenMo(
      tenChuan,
      danhSachThongSoTheoTen,
      CONFIG.traCuuTenMo,
    );
    if (theoTenMo) {
      return { thongSo: theoTenMo, nguon: "ten-mo" };
    }

    return { nguon: "khong-co" };
  };

  const hangDau = CONFIG.dongBatDauDuLieu;
  const dataNguon = sheetDich
    .getRange(`${cotNguon.sku}${hangDau}:${cotNguon.soLuongNhap}${lastRow}`)
    .getValues() as (string | number | boolean)[][];

  const dataChiSoHienTai = sheetDich
    .getRange(
      `${cotChiSo.trongLuong}${hangDau}:${cotChiSo.tongTheTich}${lastRow}`,
    )
    .getValues() as (string | number | boolean)[][];

  const dataChiSoMoi: (string | number | boolean)[][] = [];
  let soDongBoSung = 0;
  let soDongDaCoChiSo = 0;
  let soDongThieuThongSo = 0;
  let soDongKhopTenMo = 0;

  const chiSoDayDu = (
    l: string | number | boolean,
    n: string | number | boolean,
    o: string | number | boolean,
    p: string | number | boolean,
    q: string | number | boolean,
  ): boolean =>
    !oTrong(l) && !oTrong(n) && !oTrong(o) && !oTrong(p) && !oTrong(q);

  for (let i = 0; i < dataNguon.length; i++) {
    const sku = dataNguon[i][0];
    const tenHang = dataNguon[i][1];
    const hienTai = dataChiSoHienTai[i];
    const [cuL, cuM, cuN, cuO, cuP, cuQ, cuR] = hienTai;

    if (chiSoDayDu(cuL, cuN, cuO, cuP, cuQ)) {
      soDongDaCoChiSo++;
      dataChiSoMoi.push([cuL, cuM, cuN, cuO, cuP, cuQ, cuR]);
      continue;
    }

    const canBoSung =
      oTrong(cuL) ||
      oTrong(cuN) ||
      oTrong(cuO) ||
      oTrong(cuP) ||
      oTrong(cuQ) ||
      oTrong(cuM) ||
      oTrong(cuR);
    if (!canBoSung) {
      dataChiSoMoi.push([cuL, cuM, cuN, cuO, cuP, cuQ, cuR]);
      continue;
    }

    const ketQuaTraCuu = layThongSo(sku, tenHang);
    const thongSo = ketQuaTraCuu.thongSo;

    if (!thongSo) {
      if (lamSach(sku) || chuanHoaTenSanPham(tenHang)) soDongThieuThongSo++;
      dataChiSoMoi.push([cuL, cuM, cuN, cuO, cuP, cuQ, cuR]);
      continue;
    }

    const soDongSheet = hangDau + i;
    const colTL = cotChiSo.trongLuong;
    const colSL = cotNguon.soLuongNhap;
    const colTT = cotChiSo.theTich;

    const trongLuong = oTrong(cuL) ? thongSo.trongLuong || "" : cuL;
    const dai = oTrong(cuN) ? thongSo.dai || "" : cuN;
    const rong = oTrong(cuO) ? thongSo.rong || "" : cuO;
    const cao = oTrong(cuP) ? thongSo.cao || "" : cuP;
    const theTich = oTrong(cuQ) ? thongSo.theTich || "" : cuQ;

    const tongTrongLuong = oTrong(cuM)
      ? !oTrong(trongLuong)
        ? `=${colTL}${soDongSheet}*${colSL}${soDongSheet}`
        : cuM
      : cuM;

    const tongTheTich = oTrong(cuR)
      ? !oTrong(theTich)
        ? `=${colTT}${soDongSheet}*${colSL}${soDongSheet}`
        : cuR
      : cuR;

    const coThayDoi =
      trongLuong !== cuL ||
      dai !== cuN ||
      rong !== cuO ||
      cao !== cuP ||
      theTich !== cuQ ||
      tongTrongLuong !== cuM ||
      tongTheTich !== cuR;

    dataChiSoMoi.push([
      trongLuong,
      tongTrongLuong,
      dai,
      rong,
      cao,
      theTich,
      tongTheTich,
    ]);
    if (coThayDoi) {
      soDongBoSung++;
      if (ketQuaTraCuu.nguon === "ten-mo") soDongKhopTenMo++;
    }
  }

  // Khi có SKU thêm vào sau khi Auto_Nami chạy, các công thức SUBTOTAL (J1/N1/S1)
  // vẫn trỏ tới dòng cuối cũ — cập nhật lại.
  const capNhatCongThucSubtotal = (): number => {
    const cfg = CONFIG.congThucTongHop;
    if (!cfg.bat) return 0;
    let soOCapNhat = 0;
    for (const diaChi of cfg.oSubtotal) {
      const o = sheetDich.getRange(diaChi);
      const congThucMoi = namiSupportCapNhatCongThucSubtotal(
        o.getFormula(),
        lastRow,
      );
      if (congThucMoi) {
        o.setFormula(congThucMoi);
        soOCapNhat++;
      }
    }
    return soOCapNhat;
  };

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
  const apDungDinhDangTieuDeThongKe = (rangeAddress: string) => {
    const format = sheetDich.getRange(rangeAddress).getFormat();
    format.getFont().setName(CONFIG.format.fontName);
    format.getFont().setSize(CONFIG.format.fontSize);
    format.getFont().setBold(true);
    format.getFill().setColor(CONFIG.format.subtotalFillColor);
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  };
  const apDungDinhDangHeaderThongKe = (rangeAddress: string) => {
    const format = sheetDich.getRange(rangeAddress).getFormat();
    const font = format.getFont();
    font.setName(CONFIG.format.fontName);
    font.setSize(CONFIG.format.fontSize);
    font.setBold(true);
    font.setColor("white");
    format.getFill().setColor(CONFIG.format.headerFillColor);
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  };

  // Dựng lại hai bảng thống kê (giống ghiThongKeKetQua của Auto_Nami) theo dữ
  // liệu D–E hiện tại: mã phân loại / thị trường mới sẽ có dòng thống kê riêng.
  const ghiLaiThongKeKetQua = (): boolean => {
    const cfg = CONFIG.thongKeKetQua;
    if (!cfg.bat || lastRow < CONFIG.dongBatDauDuLieu) return false;

    const duLieuPhanLoai = sheetDich
      .getRange(
        `${cotNguon.thiTruong}${CONFIG.dongBatDauDuLieu}:${cotNguon.maPhanLoai}${lastRow}`,
      )
      .getValues() as (string | number | boolean)[][];
    const dsCap: NamiSupportCapThiTruongMa[] = duLieuPhanLoai.map((row) => ({
      thiTruong: String(row[0] ?? "").trim(),
      maPhanLoai: String(row[1] ?? "").trim(),
    }));

    const bangThongKe = namiSupportTaoBangThongKe(
      dsCap,
      lastRow,
      CONFIG.dongBatDauDuLieu,
      cfg.cotBatDau,
      cfg.heSoChiaTheTich,
      cfg.thuTuMaPhanLoai,
      cfg.dongBatDauBang,
    );
    if (!bangThongKe) return false;

    const layCotThongKe = (offset: number): string =>
      namiSupportDoiSoSangCot(namiSupportDoiCotSangSo(cfg.cotBatDau) + offset);
    const cot0 = layCotThongKe(0);
    const cot1 = layCotThongKe(1);
    const cot2 = layCotThongKe(2);
    const cot3 = layCotThongKe(3);
    const cot4 = layCotThongKe(4);
    const cot5 = layCotThongKe(5);

    // Xóa vùng thống kê cũ (có thể dài hơn bảng mới) trước khi ghi lại.
    let dongCuoiThongKeCu = 0;
    for (let i = 0; i < cfg.soCot; i++) {
      const dongCuoiCot = timDongCuoiTheoCot(layCotThongKe(i));
      if (dongCuoiCot > dongCuoiThongKeCu) dongCuoiThongKeCu = dongCuoiCot;
    }
    const {
      thongKeMaStartRow,
      thongKeMaRows,
      soMaPhanLoai,
      thongKeThiTruongStartRow,
      thongKeThiTruongRows,
      dsThiTruongMaPhanLoai,
    } = bangThongKe;
    const thongKeThiTruongLastRow =
      thongKeThiTruongStartRow + thongKeThiTruongRows.length - 1;
    const dongCuoiCanXoa = Math.max(dongCuoiThongKeCu, thongKeThiTruongLastRow);
    // Xóa từ hàng 1 để dọn cả bảng thống kê layout cũ (tiêu đề từng nằm ở hàng 1).
    const vungCu = sheetDich.getRange(`${cot0}1:${cot5}${dongCuoiCanXoa}`);
    vungCu.unmerge();
    vungCu.clear();

    const thongKeMaLastRow = thongKeMaStartRow + thongKeMaRows.length - 1;
    sheetDich
      .getRange(`${cot0}${thongKeMaStartRow}:${cot3}${thongKeMaLastRow}`)
      .setValues(thongKeMaRows);
    apDungDinhDangTieuDeThongKe(
      `${cot0}${thongKeMaStartRow}:${cot3}${thongKeMaStartRow}`,
    );
    apDungDinhDangHeaderThongKe(
      `${cot0}${thongKeMaStartRow + 1}:${cot3}${thongKeMaStartRow + 1}`,
    );
    apDungVien(
      sheetDich.getRange(
        `${cot0}${thongKeMaStartRow}:${cot3}${thongKeMaLastRow}`,
      ),
    );
    if (soMaPhanLoai > 0) {
      const dataStartRow = thongKeMaStartRow + 2;
      const dataRange = sheetDich.getRange(
        `${cot0}${dataStartRow}:${cot3}${thongKeMaLastRow}`,
      );
      dataRange.getFormat().getFont().setName(CONFIG.format.fontName);
      dataRange.getFormat().getFont().setSize(CONFIG.format.fontSize);
      dataRange
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      sheetDich
        .getRange(`${cot1}${dataStartRow}:${cot1}${thongKeMaLastRow}`)
        .setNumberFormat("0");
      sheetDich
        .getRange(`${cot2}${dataStartRow}:${cot3}${thongKeMaLastRow}`)
        .setNumberFormat("0.000");
    }

    sheetDich
      .getRange(
        `${cot0}${thongKeThiTruongStartRow}:${cot5}${thongKeThiTruongLastRow}`,
      )
      .setValues(thongKeThiTruongRows);
    apDungDinhDangTieuDeThongKe(
      `${cot0}${thongKeThiTruongStartRow}:${cot5}${thongKeThiTruongStartRow}`,
    );
    apDungDinhDangHeaderThongKe(
      `${cot0}${thongKeThiTruongStartRow + 1}:${cot5}${
        thongKeThiTruongStartRow + 1
      }`,
    );
    apDungVien(
      sheetDich.getRange(
        `${cot0}${thongKeThiTruongStartRow}:${cot5}${thongKeThiTruongLastRow}`,
      ),
    );
    if (dsThiTruongMaPhanLoai.length > 0) {
      const dataStartRow = thongKeThiTruongStartRow + 2;
      const dataRange = sheetDich.getRange(
        `${cot0}${dataStartRow}:${cot5}${thongKeThiTruongLastRow}`,
      );
      dataRange.getFormat().getFont().setName(CONFIG.format.fontName);
      dataRange.getFormat().getFont().setSize(CONFIG.format.fontSize);
      dataRange
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      sheetDich
        .getRange(`${cot2}${dataStartRow}:${cot2}${thongKeThiTruongLastRow}`)
        .setNumberFormat("0");
      sheetDich
        .getRange(`${cot3}${dataStartRow}:${cot5}${thongKeThiTruongLastRow}`)
        .setNumberFormat("0.000");

      let groupStartRow = dataStartRow;
      for (let i = 1; i <= dsThiTruongMaPhanLoai.length; i++) {
        const prev = dsThiTruongMaPhanLoai[i - 1];
        const current = dsThiTruongMaPhanLoai[i];
        if (current && current.thiTruong === prev.thiTruong) continue;

        const groupEndRow = dataStartRow + i - 1;
        if (groupEndRow > groupStartRow) {
          const marketRange = sheetDich.getRange(
            `${cot0}${groupStartRow}:${cot0}${groupEndRow}`,
          );
          marketRange.merge(false);
          marketRange
            .getFormat()
            .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        }
        groupStartRow = groupEndRow + 1;
      }
    }

    console.log(
      `Đã dựng lại bảng thống kê (${cot0}–${cot5}): ${soMaPhanLoai} mã phân loại, ${dsThiTruongMaPhanLoai.length} cặp thị trường–mã phân loại, dòng cuối dữ liệu ${lastRow}.`,
    );
    return true;
  };

  const soOCongThucCapNhat = capNhatCongThucSubtotal();
  if (soOCongThucCapNhat > 0) {
    console.log(
      `Đã cập nhật dòng cuối (${lastRow}) cho ${soOCongThucCapNhat} công thức SUBTOTAL của Auto_Nami.`,
    );
  }
  ghiLaiThongKeKetQua();

  if (soDongBoSung === 0) {
    console.log(
      `Không có dòng nào cần bổ sung thông số (${soDongDaCoChiSo} dòng đã có chỉ số, ${soDongThieuThongSo} dòng vẫn thiếu tra cứu).`,
    );
    return;
  }

  sheetDich
    .getRange(
      `${cotChiSo.trongLuong}${hangDau}:${cotChiSo.tongTheTich}${lastRow}`,
    )
    .setValues(dataChiSoMoi);

  sheetDich
    .getRange(
      `${cotChiSo.trongLuong}${hangDau}:${cotChiSo.cao}${lastRow}`,
    )
    .setNumberFormat("General");
  sheetDich
    .getRange(
      `${cotChiSo.theTich}${hangDau}:${cotChiSo.tongTheTich}${lastRow}`,
    )
    .setNumberFormat("0.000");

  console.log(
    `Đã bổ sung thông số (M–S) cho ${soDongBoSung} dòng. ${soDongDaCoChiSo} dòng giữ nguyên (Auto_Nami đã điền). ${soDongThieuThongSo} dòng chưa tìm thấy trong "${CONFIG.tenSheetThongSo}".`,
  );
  if (CONFIG.traCuuTenMo.bat) {
    console.log(`Khớp theo tên mờ: ${soDongKhopTenMo} dòng.`);
  }
}
