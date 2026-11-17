/**
 * ===================================================================
 * Auto_Sanji – Thống kê tồn kho theo thị trường từ sheet "Thành phẩm".
 * ===================================================================
 *
 * - Dùng cấu trúc cột nguồn tương thích Auto_Mitsuha.
 * - Tồn kho luôn tự tính từ cột nhập/xuất (không phụ thuộc cột tồn trên bảng).
 * - Có thống kê theo thị trường và theo ký tự phân loại từng thị trường
 *   (F, A, S, C, T, M... có thể cấu hình).
 * - Bảng chi tiết giữ layout cột thông số tương tự Auto_Nami.
 */

interface SanjiProductInfo {
  trongLuong: number;
  dai: number;
  rong: number;
  cao: number;
  theTich: number;
}

type SanjiThongSoTheoTenEntry = {
  tenChuan: string;
  tokens: string[];
  thongSo: SanjiProductInfo;
  coDuLieu: boolean;
};

type SanjiTraCuuTenMoConfig = {
  bat: boolean;
  diemToiThieu: number;
  tyLeTokenKhopToiThieu: number;
  soTokenKhopToiDa: number;
};

const SANJI_CONFIG = {
  tenSheetThongSo: "2026",
  xoaDuLieuCu: true,
  locThiTruong: [] as string[], // Ví dụ: ["EU", "AU", "JP"], để [] để lấy tất cả.
  traCuuTenMo: {
    bat: true,
    diemToiThieu: 0.55,
    tyLeTokenKhopToiThieu: 0.5,
    soTokenKhopToiDa: 3,
  } as SanjiTraCuuTenMoConfig,
  thongKeKyTuMacDinh: ["F", "A", "S", "C", "T", "M"],
  // Tùy biến theo từng thị trường. Ví dụ:
  // - JP: JPT4001-2 => ký tự phân loại "T"
  // - AU: AUF => "F"
  // - EU: EUM => "M"
  thongKeKyTuTheoThiTruong: {
    JP: ["F", "A", "S", "C", "T", "M"],
    AU: ["F", "A", "S", "C", "T", "M"],
    EU: ["F", "A", "S", "C", "T", "M"],
  } as { [thiTruong: string]: string[] },
  cot: {
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
      tenSheet: "Tồn kho",
    },
  },
};

// @ts-ignore: Office Scripts
async function main(workbook: ExcelScript.Workbook) {
  const CFG = SANJI_CONFIG;
  const COT_NGUON = CFG.cot.nguon;
  const COT_DICH = CFG.cot.dich;

  const sourceSheet = workbook.getWorksheet(COT_NGUON.tenSheet);
  if (!sourceSheet) {
    console.log(`LỖI: Không tìm thấy sheet nguồn "${COT_NGUON.tenSheet}".`);
    return;
  }

  const lamSach = (giaTri: string | number | boolean): string =>
    String(giaTri)
      .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
      .trim();

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
  const toSoNguyen = (n: number): number => Math.round(n);

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
    config: SanjiTraCuuTenMoConfig,
  ): number =>
    Math.max(
      config.soTokenKhopToiDa,
      Math.ceil(soTokenTraCuu * config.tyLeTokenKhopToiThieu),
    );

  const thongSoTenCoDuLieu = (thongSo: SanjiProductInfo): boolean =>
    !!(
      thongSo.trongLuong ||
      thongSo.dai ||
      thongSo.rong ||
      thongSo.cao ||
      thongSo.theTich
    );

  const timThongSoTheoTenMo = (
    tenChuan: string,
    danhSach: SanjiThongSoTheoTenEntry[],
    config: SanjiTraCuuTenMoConfig,
  ): SanjiProductInfo | undefined => {
    const queryTokens = tachTokenTenSanPham(tenChuan);
    if (queryTokens.length === 0 || danhSach.length === 0) return undefined;

    const nguongGiao = layNguongTokenKhop(queryTokens.length, config);
    let best: SanjiProductInfo | undefined;
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

  const sourceLastRow =
    sourceSheet
      .getRange(`${COT_NGUON.maHang}1048576`)
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  if (sourceLastRow < COT_NGUON.dongBatDau) {
    console.log("Sheet nguồn không có dữ liệu.");
    return;
  }

  const getColData = (col: string): (string | number | boolean)[][] => {
    if (!col) return Array(sourceLastRow - COT_NGUON.dongBatDau + 1).fill([""]);
    return sourceSheet
      .getRange(`${col}${COT_NGUON.dongBatDau}:${col}${sourceLastRow}`)
      .getValues() as (string | number | boolean)[][];
  };

  const valMaHang = getColData(COT_NGUON.maHang);
  const valTenHang = getColData(COT_NGUON.tenHang);
  const valHopDong = getColData(COT_NGUON.hopDong);
  const valSlDonHang = getColData(COT_NGUON.slDonHang);
  const valViTri = getColData(COT_NGUON.viTri);
  const valNhapKhos = COT_NGUON.cotsNhapKho.map((col) => getColData(col));
  const valXuatKhos = COT_NGUON.cotsXuatKho.map((col) => getColData(col));

  const lookupThongSo = new Map<string, SanjiProductInfo>();
  const lookupThongSoTheoTen = new Map<string, SanjiProductInfo>();
  const danhSachThongSoTheoTen: SanjiThongSoTheoTenEntry[] = [];
  const sheetThongSo = workbook.getWorksheet(CFG.tenSheetThongSo);
  if (sheetThongSo) {
    const lastRowThongSo =
      sheetThongSo
        .getRange("B1048576")
        .getRangeEdge(ExcelScript.KeyboardDirection.up)
        .getRowIndex() + 1;
    if (lastRowThongSo >= 2) {
      const dataThongSo = sheetThongSo
        .getRange(`B2:H${lastRowThongSo}`)
        .getValues() as (string | number | boolean)[][];

      for (let row of dataThongSo) {
        const maHangThongSo = lamSach(row[0]);
        const tenHangThongSo = chuanHoaTenSanPham(row[1]);
        if (!maHangThongSo && !tenHangThongSo) continue;

        const trongLuong = parseSoTuExcel(row[3]);
        const dai = parseSoTuExcel(row[4]);
        const rong = parseSoTuExcel(row[5]);
        const caoGoc = parseSoTuExcel(row[6]);
        const cao = caoGoc > 0 ? caoGoc : dai > 0 && rong > 0 ? 1.1 : 0;
        const theTich = (dai * rong * cao) / 1000000;

        const thongSo: SanjiProductInfo = {
          trongLuong: trongLuong ? Number(trongLuong.toFixed(4)) : 0,
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
            const idx = danhSachThongSoTheoTen.findIndex(
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
      `Lưu ý: Không tìm thấy sheet "${CFG.tenSheetThongSo}". Cột thông số sẽ để trống.`,
    );
  }

  const layThongSoHang = (
    maHang: string | number | boolean,
    tenHang: string | number | boolean,
  ): SanjiProductInfo | undefined => {
    const ma = lamSach(maHang);
    if (ma) {
      const theoMa = lookupThongSo.get(ma);
      if (theoMa) return theoMa;
    }
    const tenChuan = chuanHoaTenSanPham(tenHang);
    if (!tenChuan) return undefined;

    const theoTenChinhXac = lookupThongSoTheoTen.get(tenChuan);
    if (theoTenChinhXac) return theoTenChinhXac;
    if (!CFG.traCuuTenMo.bat || danhSachThongSoTheoTen.length === 0) {
      return undefined;
    }
    return timThongSoTheoTenMo(
      tenChuan,
      danhSachThongSoTheoTen,
      CFG.traCuuTenMo,
    );
  };

  const chuanHoaDanhSachKyTu = (list: string[]): string[] => {
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
  };

  const dsKyTuMacDinh = chuanHoaDanhSachKyTu(CFG.thongKeKyTuMacDinh);
  const layDanhSachKyTuTheoThiTruong = (thiTruong: string): string[] => {
    const dsCauHinh = CFG.thongKeKyTuTheoThiTruong[thiTruong];
    if (!dsCauHinh || dsCauHinh.length === 0) return dsKyTuMacDinh;
    return chuanHoaDanhSachKyTu(dsCauHinh);
  };

  const layThiTruong = (hopDong: string): string => {
    const hd = String(hopDong ?? "")
      .trim()
      .toUpperCase();
    const match = hd.match(/^([A-Z]{2})/);
    return match ? match[1] : "KHAC";
  };

  const layKyTuPhanLoai = (hopDong: string, thiTruong: string): string => {
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

  type SanjiTongHopThiTruong = {
    soSKU: number;
    tongTon: number;
    tongTrongLuong: number;
    tongTheTich: number;
    tonTheoKyTu: { [k: string]: number };
  };

  let detailRows: (string | number | boolean)[][] = [];
  let tongHopThiTruong: { [thiTruong: string]: SanjiTongHopThiTruong } = {};
  let stt = 1;
  let tapKyTuTongHop = new Set<string>(dsKyTuMacDinh);
  const dsLocThiTruong = (CFG.locThiTruong || []).map((x) =>
    String(x).trim().toUpperCase(),
  );
  const dangLocThiTruong = dsLocThiTruong.length > 0;

  for (let i = 0; i < valMaHang.length; i++) {
    const maHang = lamSach(valMaHang[i][0]);
    if (!maHang) continue;

    const hopDong = String(valHopDong[i][0] ?? "").trim();
    if (!hopDong) continue;
    const thiTruong = layThiTruong(hopDong);
    if (dangLocThiTruong && !dsLocThiTruong.includes(thiTruong)) continue;

    const dsKyTuThiTruong = layDanhSachKyTuTheoThiTruong(thiTruong);
    dsKyTuThiTruong.forEach((k) => tapKyTuTongHop.add(k));

    const kyTuPhanLoaiRaw = layKyTuPhanLoai(hopDong, thiTruong);
    const kyTuPhanLoai = dsKyTuThiTruong.includes(kyTuPhanLoaiRaw)
      ? kyTuPhanLoaiRaw
      : "KHAC";

    const tenHang = String(valTenHang[i][0] ?? "").trim();
    const slDonHang = toSoNguyen(parseSoTuExcel(valSlDonHang[i][0]));

    let tongNhap = 0;
    for (let dataNhap of valNhapKhos) {
      tongNhap += parseSoTuExcel(dataNhap[i][0]);
    }
    tongNhap = toSoNguyen(tongNhap);

    let tongXuat = 0;
    for (let dataXuat of valXuatKhos) {
      tongXuat += parseSoTuExcel(dataXuat[i][0]);
    }
    tongXuat = toSoNguyen(tongXuat);

    let tonTuTinh = toSoNguyen(tongNhap - tongXuat);
    if (tonTuTinh <= 0) continue;

    const thongSoHang = layThongSoHang(maHang, tenHang);
    const trongLuong = thongSoHang?.trongLuong || "";
    const dai = thongSoHang?.dai || "";
    const rong = thongSoHang?.rong || "";
    const cao = thongSoHang?.cao || "";
    const theTich = thongSoHang?.theTich || "";
    const tongTrongLuong =
      typeof trongLuong === "number"
        ? Number((trongLuong * tonTuTinh).toFixed(4))
        : "";
    const tongTheTich =
      typeof theTich === "number"
        ? Number((theTich * tonTuTinh).toFixed(3))
        : "";

    detailRows.push([
      stt,
      thiTruong,
      kyTuPhanLoai,
      hopDong,
      maHang,
      tenHang,
      String(valViTri[i][0] ?? "").trim(),
      slDonHang,
      tongNhap,
      tongXuat,
      tonTuTinh,
      trongLuong,
      tongTrongLuong,
      dai,
      rong,
      cao,
      theTich,
      tongTheTich,
    ]);
    stt++;

    if (!tongHopThiTruong[thiTruong]) {
      tongHopThiTruong[thiTruong] = {
        soSKU: 0,
        tongTon: 0,
        tongTrongLuong: 0,
        tongTheTich: 0,
        tonTheoKyTu: {},
      };
    }
    const agg = tongHopThiTruong[thiTruong];
    agg.soSKU += 1;
    agg.tongTon = toSoNguyen(agg.tongTon + tonTuTinh);
    agg.tongTrongLuong = Number(
      (
        agg.tongTrongLuong +
        (typeof tongTrongLuong === "number" ? tongTrongLuong : 0)
      ).toFixed(4),
    );
    agg.tongTheTich = Number(
      (
        agg.tongTheTich + (typeof tongTheTich === "number" ? tongTheTich : 0)
      ).toFixed(3),
    );
    agg.tonTheoKyTu[kyTuPhanLoai] = toSoNguyen(
      (agg.tonTheoKyTu[kyTuPhanLoai] || 0) + tonTuTinh,
    );
  }

  if (detailRows.length === 0) {
    console.log("Không có dữ liệu tồn kho dương thỏa điều kiện để thống kê.");
    return;
  }

  const dsKyTuTongHop = Array.from(tapKyTuTongHop).sort();
  const headerTongHop = [
    "Thị trường",
    "Số SKU tồn",
    "Tổng tồn",
    "Tổng trọng lượng",
    "Tổng thể tích",
    ...dsKyTuTongHop.map((k) => `SL-${k}`),
    "SL-KHAC",
  ];

  const thiTruongSorted = Object.keys(tongHopThiTruong).sort();
  let rowsTongHop: (string | number | boolean)[][] = [];

  for (let tt of thiTruongSorted) {
    const agg = tongHopThiTruong[tt];
    let row: (string | number | boolean)[] = [
      tt,
      agg.soSKU,
      agg.tongTon,
      agg.tongTrongLuong,
      agg.tongTheTich,
    ];
    for (let k of dsKyTuTongHop) {
      let v = toSoNguyen(agg.tonTheoKyTu[k] || 0);
      row.push(v);
    }
    const vKhac = toSoNguyen(agg.tonTheoKyTu.KHAC || 0);
    row.push(vKhac);
    rowsTongHop.push(row);
  }

  let rowTongCongTongHop: (string | number | boolean)[] = [
    "TỔNG CỘNG",
    "",
    "",
    "",
    "",
  ];
  for (let k of dsKyTuTongHop) {
    rowTongCongTongHop.push("");
  }
  rowTongCongTongHop.push("");

  const headerChiTiet = [
    "STT",
    "Thị trường",
    "Mã phân loại",
    "Mã hợp đồng",
    "SKU",
    "Tên hàng",
    "Vị trí",
    "Số lượng đơn hàng",
    "Tổng nhập",
    "Tổng xuất",
    "Tồn",
    "Trọng lượng",
    "Tổng trọng lượng",
    "Dài",
    "Rộng",
    "Cao",
    "Thể tích",
    "Tổng thể tích",
  ];

  const rowTongCongChiTiet: (string | number | boolean)[] = [
    "",
    "",
    "",
    "TỔNG CỘNG",
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
    "",
    "",
    "",
  ];

  let rows: (string | number | boolean)[][] = [];
  rows.push([`THỐNG KÊ TỒN TỪ "${COT_NGUON.tenSheet}"`]);
  rows.push([""]);
  const rowTongCongTongHopIndex = rows.length + 1;
  rows.push(rowTongCongTongHop);
  const rowHeaderTongHop = rows.length + 1;
  rows.push(headerTongHop);
  rows = rows.concat(rowsTongHop);
  rows.push([""]);
  const rowTongCongChiTietIndex = rows.length + 1;
  rows.push(rowTongCongChiTiet);
  const rowHeaderChiTiet = rows.length + 1;
  rows.push(headerChiTiet);
  rows = rows.concat(detailRows);

  const maxCols = rows.reduce((max, r) => Math.max(max, r.length), 1);
  const padded = rows.map((r) => {
    let x = r.slice() as (string | number | boolean)[];
    while (x.length < maxCols) x.push("");
    return x;
  });

  const colToLetter = (n: number): string => {
    let out = "";
    let x = n;
    while (x > 0) {
      const mod = (x - 1) % 26;
      out = String.fromCharCode(65 + mod) + out;
      x = Math.floor((x - 1) / 26);
    }
    return out;
  };
  const taoCongThucSubtotal = (
    cot: string,
    dongBatDau: number,
    dongKetThuc: number,
  ): string => `=SUBTOTAL(9,${cot}${dongBatDau}:${cot}${dongKetThuc})`;
  const lastCol = colToLetter(maxCols);
  const lastColTongHop = colToLetter(headerTongHop.length);

  let targetSheet = workbook.getWorksheet(COT_DICH.tenSheet);
  if (!targetSheet) targetSheet = workbook.addWorksheet(COT_DICH.tenSheet);
  if (CFG.xoaDuLieuCu)
    targetSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);

  targetSheet.getRange(`A1:${lastCol}${padded.length}`).setValues(padded);
  const rowTongHopDataStart = rowHeaderTongHop + 1;
  const rowTongHopDataEnd = rowHeaderTongHop + rowsTongHop.length;
  const congThucTongHop = Array.from(
    { length: headerTongHop.length - 1 },
    (_, i) => {
      const cot = colToLetter(i + 2);
      return taoCongThucSubtotal(cot, rowTongHopDataStart, rowTongHopDataEnd);
    },
  );
  targetSheet
    .getRange(
      `B${rowTongCongTongHopIndex}:${lastColTongHop}${rowTongCongTongHopIndex}`,
    )
    .setFormulas([congThucTongHop]);

  const rowChiTietDataStart = rowHeaderChiTiet + 1;
  const rowChiTietDataEnd = rowHeaderChiTiet + detailRows.length;
  targetSheet
    .getRange(`K${rowTongCongChiTietIndex}`)
    .setFormula(
      taoCongThucSubtotal("K", rowChiTietDataStart, rowChiTietDataEnd),
    );
  targetSheet
    .getRange(`M${rowTongCongChiTietIndex}`)
    .setFormula(
      taoCongThucSubtotal("M", rowChiTietDataStart, rowChiTietDataEnd),
    );
  targetSheet
    .getRange(`R${rowTongCongChiTietIndex}`)
    .setFormula(
      taoCongThucSubtotal("R", rowChiTietDataStart, rowChiTietDataEnd),
    );
  const autoFilter = targetSheet.getAutoFilter();
  if (autoFilter.getEnabled()) autoFilter.remove();
  autoFilter.apply(`A${rowHeaderChiTiet}:R${rowChiTietDataEnd}`);

  // Tiêu đề
  targetSheet.getRange(`A1:${lastCol}1`).merge(false);
  targetSheet.getRange("A1").getFormat().getFont().setBold(true);
  targetSheet.getRange("A1").getFormat().getFont().setSize(14);
  targetSheet
    .getRange("A1")
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Header tổng hợp + chi tiết
  targetSheet
    .getRange(`A${rowHeaderTongHop}:${lastCol}${rowHeaderTongHop}`)
    .getFormat()
    .getFont()
    .setBold(true);
  targetSheet
    .getRange(`A${rowHeaderTongHop}:${lastCol}${rowHeaderTongHop}`)
    .getFormat()
    .getFill()
    .setColor("#D9E1F2");
  targetSheet
    .getRange(`A${rowHeaderChiTiet}:${lastCol}${rowHeaderChiTiet}`)
    .getFormat()
    .getFont()
    .setBold(true);
  targetSheet
    .getRange(`A${rowHeaderChiTiet}:${lastCol}${rowHeaderChiTiet}`)
    .getFormat()
    .getFill()
    .setColor("#BDD7EE");
  targetSheet
    .getRange(
      `A${rowTongCongTongHopIndex}:${lastCol}${rowTongCongTongHopIndex}`,
    )
    .getFormat()
    .getFont()
    .setBold(true);
  targetSheet
    .getRange(
      `A${rowTongCongChiTietIndex}:${lastCol}${rowTongCongChiTietIndex}`,
    )
    .getFormat()
    .getFont()
    .setBold(true);

  // Căn giữa toàn bộ và căn trái một vài cột text của bảng chi tiết
  targetSheet
    .getRange(`A1:${lastCol}${padded.length}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  const rowDetailStart = rowHeaderChiTiet + 1;
  const rowDetailEnd = rowChiTietDataEnd;
  targetSheet
    .getRange(`D${rowTongCongChiTietIndex}:G${rowDetailEnd}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  // Định dạng số
  targetSheet
    .getRange(`B${rowTongCongTongHopIndex}:B${rowTongHopDataEnd}`)
    .setNumberFormat("0");
  targetSheet
    .getRange(`C${rowTongCongTongHopIndex}:C${rowTongHopDataEnd}`)
    .setNumberFormat("0");
  targetSheet
    .getRange(`D${rowTongCongTongHopIndex}:D${rowTongHopDataEnd}`)
    .setNumberFormat("0.0000");
  targetSheet
    .getRange(`E${rowTongCongTongHopIndex}:E${rowTongHopDataEnd}`)
    .setNumberFormat("0.000");
  if (maxCols >= 6) {
    targetSheet
      .getRange(
        `F${rowTongCongTongHopIndex}:${lastColTongHop}${rowTongHopDataEnd}`,
      )
      .setNumberFormat("0");
  }
  targetSheet
    .getRange(`Q${rowTongCongChiTietIndex}:R${rowDetailEnd}`)
    .setNumberFormat("0.000");
  targetSheet
    .getRange(`L${rowTongCongChiTietIndex}:P${rowDetailEnd}`)
    .setNumberFormat("0.0000");
  targetSheet
    .getRange(`H${rowTongCongChiTietIndex}:K${rowDetailEnd}`)
    .setNumberFormat("0");

  // Viền toàn vùng
  const fullRange = targetSheet.getRange(`A1:${lastCol}${padded.length}`);
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((bIndex) => {
    fullRange
      .getFormat()
      .getRangeBorder(bIndex)
      .setStyle(ExcelScript.BorderLineStyle.continuous);
  });

  targetSheet
    .getRange(`A1:${lastCol}${padded.length}`)
    .getFormat()
    .autofitColumns();
  targetSheet.activate();

  console.log(
    `Auto_Sanji hoàn tất: ${detailRows.length} dòng chi tiết tồn, ${rowsTongHop.length} thị trường.`,
  );
}
