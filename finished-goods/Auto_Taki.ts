/**
 * ===================================================================
 *  Auto_Taki – Tự động xuất kho theo ngày, tháng, năm.
 * ===================================================================
 *
 * Script này tự động xuất kho theo ngày, tháng, năm dựa trên cấu hình.
 * Mọi thông số đều nằm trong CONFIG → dễ tùy chỉnh, mở rộng.
 *
 */
// ==========================================
// 1. CẤU HÌNH (CONFIGURATION)
// ==========================================
const TAKI_CONFIG = {
  day: "08/05/2026", // Ngày xuất kho
  tonKhoToiThieuMacDinh: 5, // Ngưỡng tồn kho tối thiểu mặc định
  tonKhoToiThieuTheoThiTruong: {
    // Lấy theo 2 ký tự đầu của hợp đồng ở cột B source. Có thể mở rộng thêm: UA, PL...
    EU: 8,
  } as Record<string, number>,
  xuatTheoCapTheoThiTruong: {
    UE: 8, // Ví dụ: UE chỉ xuất khi tổng nhóm HĐ+SKU đạt ngưỡng 8
  } as Record<string, number>,
  cot: {
    // Sheet nguồn (sheet hiện tại đang mở)
    source: {
      tenSheet: "auto",
      maHang: "C",
      hopDong: "B",
      slDaXuat: "H",
      slDongGoi: "I",
      ghiChu: "P",
      ketQua: "R",
      dongBatDau: 2,
      expectedHeaderGhiChu: "GHICHÚ备注",
    },
    // Sheet đích (sheet kho)
    target: {
      tenSheet: "成品出入库明细  汇总表",
      maHang: "D",
      hopDong: "J",
      viTri: "K",
      nhap1: "M",
      nhap2: "O",
      nhap3: "Q",
      tongNhap: "S",
      tonKho: "AJ",
      chuaHoanThanh: "R",
      slDonHang: "H",
      dongBatDau: 2,
      expectedHeaderTongNhap: "入库合计Tổngnhậpkho",
      expectedHeaderTonKho: "库存Tồnkho",
      expectedHeaderChuaHoanThanh:
        "订单未完成余数Sốlượngchưahoànthànhcủađơnhàng",
      expectedHeaderSlDonHang: "订单数量Sốlượngđơnhàng",
      expectedHeaderMaHang: "客户物料编码SKU",
      expectedHeaderHopDong: "计划跟踪号Thịtrường",
      expectedHeaderViTri: "库位Vịtrí",
      // MẢNG CẤU HÌNH CỘT XUẤT KHO (Rất linh hoạt, có thể thêm cột 4, 5... dễ dàng)
      cotsXuatKho: [
        {
          sl: "AC",
          ngay: "AB",
          exSl: "出库1Xuấtkho1",
          exNgay: "出库日期Ngàyxuấtkho",
        },
        {
          sl: "AE",
          ngay: "AD",
          exSl: "出库2Xuấtkho2",
          exNgay: "出库日期Ngàyxuấtkho",
        },
        {
          sl: "AG",
          ngay: "AF",
          exSl: "出库3Xuấtkho3",
          exNgay: "出库日期Ngàyxuấtkho",
        },
      ],
    },
  },
};

// @ts-ignore: Tránh cảnh báo trùng lặp hàm khi code nhiều file TS trong cùng thư mục
function lamSach(giaTri: string | number | boolean): string {
  return String(giaTri)
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "") // Xóa mọi loại khoảng trắng ẩn
    .trim();
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
  return lamSach(raw);
}

function ngayToSerial(dateStr: string): number {
  // Chuyển chuỗi dd/mm/yyyy sang Excel serial number
  // để tránh Excel tự phân tích theo locale US (mm/dd/yyyy)
  let parts = dateStr.split("/");
  let d = Number(parts[0]);
  let m = Number(parts[1]) - 1; // JS tháng bắt đầu từ 0
  let y = Number(parts[2]);
  let jsDate = new Date(Date.UTC(y, m, d));
  return Math.floor(jsDate.getTime() / 86400000) + 25569;
}

/** Khối dọc một cột (hàng 1-based) — dùng getRange(A1) thay getRangeByIndexes. */
function layVungCotDoc(
  sheet: ExcelScript.Worksheet,
  col: string,
  rowDau: number,
  rowCuoi: number,
): ExcelScript.Range {
  const cotChu = lamSach(col).toUpperCase();
  return sheet.getRange(`${cotChu}${rowDau}:${cotChu}${rowCuoi}`);
}

/** Ghi sheet đích theo từng khối dòng liên tiếp trên cùng cột — tránh timeout do gọi API từng ô. */
function ghiTargetUpdatesTheoLo(
  targetSheet: ExcelScript.Worksheet,
  updates: {
    row: number;
    col: string;
    val: string | number;
    isDate?: boolean;
  }[],
) {
  if (updates.length === 0) return;

  type U = (typeof updates)[0];
  let theoCot = new Map<string, U[]>();
  let vungNgayTheoCot = new Map<string, { minRow: number; maxRow: number }>();
  for (let u of updates) {
    let arr = theoCot.get(u.col);
    if (!arr) {
      arr = [];
      theoCot.set(u.col, arr);
    }
    arr.push(u);

    if (u.isDate) {
      let vung = vungNgayTheoCot.get(u.col);
      if (!vung) {
        vungNgayTheoCot.set(u.col, { minRow: u.row, maxRow: u.row });
      } else {
        vung.minRow = Math.min(vung.minRow, u.row);
        vung.maxRow = Math.max(vung.maxRow, u.row);
      }
    }
  }

  for (let [, list] of Array.from(theoCot.entries())) {
    list.sort((a: U, b: U) => a.row - b.row);
    let i = 0;
    while (i < list.length) {
      let start = i;
      let laNgay = !!list[i].isDate;
      while (
        i + 1 < list.length &&
        list[i + 1].row === list[i].row + 1 &&
        !!list[i + 1].isDate === laNgay
      ) {
        i++;
      }
      let doan = list.slice(start, i + 1);
      let rowDau = doan[0].row;
      let rowCuoi = doan[doan.length - 1].row;
      let cot = doan[0].col;
      let rng = layVungCotDoc(targetSheet, cot, rowDau, rowCuoi);
      let giaTri: (string | number | boolean)[][] = doan.map((u: U) =>
        laNgay ? [ngayToSerial(String(u.val))] : [u.val],
      );
      rng.setValues(giaTri);
      i++;
    }
  }

  for (let [cot, vung] of Array.from(vungNgayTheoCot.entries())) {
    // setNumberFormatLocal dễ timeout trên Office Scripts; format invariant vẫn hiển thị dd/mm/yyyy.
    layVungCotDoc(targetSheet, cot, vung.minRow, vung.maxRow).setNumberFormat(
      "dd/mm/yyyy",
    );
  }
}

function parseOptionalNumber(raw: string | number | boolean): number | null {
  let s = lamSach(raw);
  if (s === "") return 0;
  let n = Number(s);
  return isNaN(n) ? null : n;
}

function tachSKUCapAB(maHang: string | number | boolean): {
  laSKUCapAB: boolean;
  skuGoc: string;
  hauTo: "A" | "B" | "";
} {
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
}

// Expose pure helpers for Node-based regression tests without affecting Excel execution.
const takiTestables = {
  lamSach,
  chuanHoaNgay,
  ngayToSerial,
  ghiTargetUpdatesTheoLo,
  parseOptionalNumber,
  tachSKUCapAB,
};

(
  globalThis as typeof globalThis & {
    __AUTO_TAKI_TESTABLES__?: typeof takiTestables;
  }
).__AUTO_TAKI_TESTABLES__ = takiTestables;

// Main Function
// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
async function main(workbook: ExcelScript.Workbook) {
  const CONFIG = TAKI_CONFIG;
  const COT = CONFIG.cot;

  // Xác định ngày sử dụng
  let dayToUse = CONFIG.day;
  if (!dayToUse || dayToUse.trim() === "") {
    let now = new Date();
    let dd = String(now.getDate()).padStart(2, "0");
    let mm = String(now.getMonth() + 1).padStart(2, "0");
    let yyyy = now.getFullYear();
    dayToUse = dd + "/" + mm + "/" + yyyy;
  }
  const dayNorm = dayToUse.replace(/\D/g, ""); // Dùng để so sánh ngày
  const getTonKhoToiThieuTheoHopDong = (hopDong: string): number => {
    let prefix = lamSach(hopDong).toUpperCase().slice(0, 2);
    let byMarket = CONFIG.tonKhoToiThieuTheoThiTruong[prefix];
    if (typeof byMarket === "number" && !isNaN(byMarket) && byMarket >= 0) {
      return byMarket;
    }
    return CONFIG.tonKhoToiThieuMacDinh;
  };
  const getNguongXuatTheoCap = (hopDong: string): number => {
    let prefix = lamSach(hopDong).toUpperCase().slice(0, 2);
    let nguong = CONFIG.xuatTheoCapTheoThiTruong[prefix];
    return typeof nguong === "number" && !isNaN(nguong) && nguong > 0
      ? nguong
      : 0;
  };

  // ==========================================
  // 2. KHỞI TẠO & KIỂM TRA SHEET NGUỒN
  // ==========================================
  let selectedSheet = workbook.getActiveWorksheet();
  if (selectedSheet.getName() !== COT.source.tenSheet) {
    console.log(
      `Sai sheet! Hãy mở sheet "${COT.source.tenSheet}" rồi chạy lại.`,
    );
    return;
  }

  let bd = COT.source.dongBatDau;
  let lastRow =
    selectedSheet
      .getRange(COT.source.maHang + "1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  if (lastRow < bd) {
    console.log("Sheet hiện tại đang trống hoặc cột Mã Hàng không có dữ liệu!");
    return;
  }

  let headerGhiChu = lamSach(
    selectedSheet.getRange(COT.source.ghiChu + "1").getValue(),
  );
  if (headerGhiChu !== COT.source.expectedHeaderGhiChu) {
    console.log(
      `LỖI: Tiêu đề cột ${COT.source.ghiChu} nguồn sai. Mong đợi: '${COT.source.expectedHeaderGhiChu}'.`,
    );
    return;
  }

  // Đọc dữ liệu nguồn
  let slDaXuatNguon = selectedSheet
    .getRange(COT.source.slDaXuat + bd + ":" + COT.source.slDaXuat + lastRow)
    .getValues();
  let sldonggoi = selectedSheet
    .getRange(COT.source.slDongGoi + bd + ":" + COT.source.slDongGoi + lastRow)
    .getValues();
  let mahang = selectedSheet
    .getRange(COT.source.maHang + bd + ":" + COT.source.maHang + lastRow)
    .getValues();
  let sohopdong = selectedSheet
    .getRange(COT.source.hopDong + bd + ":" + COT.source.hopDong + lastRow)
    .getValues();
  let ghiChuSource = selectedSheet
    .getRange(COT.source.ghiChu + bd + ":" + COT.source.ghiChu + lastRow)
    .getValues();

  // Tiền kiểm tra (Pre-validate) TOÀN BỘ dữ liệu đầu vào để bảo đảm tính toàn vẹn
  let checkTrungLap = new Set<string>();
  let thongKeCapABNguon = new Map<string, { tongA: number; tongB: number }>();
  let tongXuatTheoNhomNguon = new Map<string, number>();
  let dataError = "";
  for (let k = 0; k < mahang.length; k++) {
    let valMH = String(mahang[k][0]).trim().toUpperCase();
    if (valMH === "") continue;

    // 1. Kiểm tra định dạng mã hàng (từ 6 chữ số trở lên, có thể kèm A/B)
    if (!/^\d{6,}([AB])?$/.test(valMH)) {
      dataError = `LỖI TỪ CHỐI: Dòng ${k + bd} có mã hàng không hợp lệ (yêu cầu từ 6 chữ số trở lên, có thể kèm A/B): '${valMH}'`;
      break;
    }

    // 2. Kiểm tra tính hợp lệ của số lượng xuất (không được bỏ trống, bằng 0, số âm, hoặc chữ cái)
    let valSLStr = lamSach(sldonggoi[k][0]);
    let valSL = Number(valSLStr);
    if (valSLStr === "" || isNaN(valSL) || valSL <= 0) {
      dataError = `LỖI TỪ CHỐI: Dòng ${k + bd} có số lượng xuất không hợp lệ (phải là số > 0): '${valSLStr}'`;
      break;
    }

    // 2.1 slDaXuat (nguồn) nếu có nhập thì phải là số >= 0
    let valSLDaXuatStr = lamSach(slDaXuatNguon[k][0]);
    if (valSLDaXuatStr !== "") {
      let valSLDaXuat = Number(valSLDaXuatStr);
      if (isNaN(valSLDaXuat) || valSLDaXuat < 0) {
        dataError = `LỖI TỪ CHỐI: Dòng ${k + bd} có SL đã xuất không hợp lệ (phải là số >= 0): '${valSLDaXuatStr}'`;
        break;
      }
    }

    // 3. Kiểm tra chống xuất lặp lại cùng một mặt hàng trong cùng 1 sheet nguồn
    let valHD = String(sohopdong[k][0]).trim().toUpperCase();
    let valGC = lamSach(ghiChuSource[k][0]);
    let khoaCheck = `${valMH}|${valHD}|${valGC}`;
    if (checkTrungLap.has(khoaCheck)) {
      dataError = `LỖI TỪ CHỐI: Phát hiện dòng xuất lặp lại ở sheet nguồn!\nMã: [${valMH}] - HĐ: [${valHD}] - Vị trí: [${valGC}]`;
      break;
    }
    checkTrungLap.add(khoaCheck);

    // 4. Thống kê SKU A/B theo nhóm thị trường + hợp đồng + SKU gốc
    // để chặn trường hợp chỉ xuất một vế A hoặc B.
    let thongTinCapAB = tachSKUCapAB(valMH);
    if (thongTinCapAB.laSKUCapAB) {
      let thiTruong = lamSach(valHD).toUpperCase().slice(0, 2);
      let khoaCapAB = `${thiTruong}|${valHD}|${thongTinCapAB.skuGoc}`;
      if (!thongKeCapABNguon.has(khoaCapAB)) {
        thongKeCapABNguon.set(khoaCapAB, { tongA: 0, tongB: 0 });
      }
      let thongKe = thongKeCapABNguon.get(khoaCapAB)!;
      if (thongTinCapAB.hauTo === "A") {
        thongKe.tongA = Number((thongKe.tongA + valSL).toFixed(4));
      } else if (thongTinCapAB.hauTo === "B") {
        thongKe.tongB = Number((thongKe.tongB + valSL).toFixed(4));
      }
    }

    // 5. Thống kê tổng xuất theo nhóm thị trường + HĐ + SKU
    // để áp quy tắc ngưỡng xuất theo cặp theo thị trường ngay tại Taki.
    let thiTruong = lamSach(valHD).toUpperCase().slice(0, 2);
    let khoaNhom = `${thiTruong}|${valHD}|${valMH}`;
    tongXuatTheoNhomNguon.set(
      khoaNhom,
      Number(((tongXuatTheoNhomNguon.get(khoaNhom) || 0) + valSL).toFixed(4)),
    );
  }

  if (dataError !== "") {
    console.log(dataError);
    return;
  }

  let ketQuaRange = selectedSheet.getRange(
    COT.source.ketQua + bd + ":" + COT.source.ketQua + lastRow,
  );
  ketQuaRange.clear();

  // Chỉ tự bổ sung công thức Tổng nhập cho những SKU có xuất hiện ở sheet nguồn.
  let sourceSkuSet = new Set<string>();
  for (let i = 0; i < mahang.length; i++) {
    let sku = lamSach(mahang[i][0]).toUpperCase();
    if (sku !== "") sourceSkuSet.add(sku);
  }

  // ==========================================
  // 3. KHỞI TẠO & KIỂM TRA SHEET ĐÍCH
  // ==========================================
  let targetSheet = workbook.getWorksheet(COT.target.tenSheet);
  if (!targetSheet) {
    console.log(`Không tìm thấy Sheet: ${COT.target.tenSheet}`);
    return;
  }

  let autoFilter = targetSheet.getAutoFilter();
  if (autoFilter) autoFilter.remove();

  let targetLastRow =
    targetSheet
      .getRange(COT.target.maHang + "1048576")
      .getRangeEdge(ExcelScript.KeyboardDirection.up)
      .getRowIndex() + 1;
  if (targetLastRow < COT.target.dongBatDau) {
    console.log("Sheet kho đang trống!");
    return;
  }

  // Khử lặp (DRY) trong việc validate tiêu đề đích
  const headersToValidate = [
    { col: COT.target.maHang, expected: COT.target.expectedHeaderMaHang },
    { col: COT.target.hopDong, expected: COT.target.expectedHeaderHopDong },
    { col: COT.target.viTri, expected: COT.target.expectedHeaderViTri },
    { col: COT.target.tongNhap, expected: COT.target.expectedHeaderTongNhap },
    { col: COT.target.tonKho, expected: COT.target.expectedHeaderTonKho },
    {
      col: COT.target.chuaHoanThanh,
      expected: COT.target.expectedHeaderChuaHoanThanh,
    },
    { col: COT.target.slDonHang, expected: COT.target.expectedHeaderSlDonHang },
  ];
  COT.target.cotsXuatKho.forEach((c) => {
    headersToValidate.push(
      { col: c.sl, expected: c.exSl },
      { col: c.ngay, expected: c.exNgay },
    );
  });

  // Đọc toàn bộ dòng 1 (A1:CZ1) một lần duy nhất để tránh gọi API getRange trong vòng lặp
  let headerRowVals = targetSheet.getRange("A1:CZ1").getValues()[0];
  let colToIndex = (col: string) => {
    let idx = 0;
    for (let i = 0; i < col.length; i++) {
      idx = idx * 26 + col.charCodeAt(i) - 64;
    }
    return idx - 1;
  };

  let headerError = "";
  for (let hc of headersToValidate) {
    let colIdx = colToIndex(hc.col);
    let headerVal = lamSach(headerRowVals[colIdx]);
    if (headerVal !== hc.expected) {
      headerError = `LỖI: Tiêu đề cột ${hc.col} đích sai. Mong đợi: '${hc.expected}', Thực tế: '${headerVal}'.`;
      break;
    }
  }

  if (headerError !== "") {
    console.log(headerError);
    return;
  }

  // ==========================================
  // 4. ĐỌC DỮ LIỆU ĐÍCH VÀ TẠO TỪ ĐIỂN
  // ==========================================
  let tbd = COT.target.dongBatDau;
  let targetColMH = targetSheet
    .getRange(COT.target.maHang + tbd + ":" + COT.target.maHang + targetLastRow)
    .getValues();
  let targetColHD = targetSheet
    .getRange(
      COT.target.hopDong + tbd + ":" + COT.target.hopDong + targetLastRow,
    )
    .getValues();
  let targetColViTri = targetSheet
    .getRange(COT.target.viTri + tbd + ":" + COT.target.viTri + targetLastRow)
    .getValues();

  // Đảm bảo cột Tổng nhập (S) có công thức nếu bị trống.
  // Ví dụ: S123 = M123 + O123 + Q123
  let targetRangeTongNhap = targetSheet.getRange(
    COT.target.tongNhap + tbd + ":" + COT.target.tongNhap + targetLastRow,
  );
  let targetColTongNhapFormulas = targetRangeTongNhap.getFormulas();
  let targetColTongNhapExistingValues = targetRangeTongNhap.getValues();
  let diaChiCanThemCongThucTongNhap: { rowNum: number; formula: string }[] = [];
  for (let j = 0; j < targetColTongNhapFormulas.length; j++) {
    let rowNum = j + tbd;
    let targetSku = lamSach(targetColMH[j][0]).toUpperCase();
    let currentFormula = lamSach(targetColTongNhapFormulas[j][0]);
    let currentTongNhapValue = lamSach(targetColTongNhapExistingValues[j][0]);
    if (
      sourceSkuSet.has(targetSku) &&
      currentFormula === "" &&
      currentTongNhapValue === ""
    ) {
      diaChiCanThemCongThucTongNhap.push({
        rowNum,
        formula: `=${COT.target.nhap1}${rowNum}+${COT.target.nhap2}${rowNum}+${COT.target.nhap3}${rowNum}`,
      });
    }
  }
  if (diaChiCanThemCongThucTongNhap.length > 0) {
    for (let item of diaChiCanThemCongThucTongNhap) {
      targetSheet
        .getRange(`${COT.target.tongNhap}${item.rowNum}`)
        .setFormula(item.formula);
    }
  }

  // Đảm bảo cột Tồn kho (AJ) có công thức nếu bị trống.
  // Ví dụ: AJ123 = S123-AC123-AE123-AG123 (theo mảng cấu hình cột xuất kho).
  let targetRangeTonKho = targetSheet.getRange(
    COT.target.tonKho + tbd + ":" + COT.target.tonKho + targetLastRow,
  );
  let targetColTonKhoFormulas = targetRangeTonKho.getFormulas();
  let targetColTonKhoExistingValues = targetRangeTonKho.getValues();
  let diaChiCanThemCongThucTonKho: { rowNum: number; formula: string }[] = [];
  for (let j = 0; j < targetColTonKhoFormulas.length; j++) {
    let rowNum = j + tbd;
    let targetSku = lamSach(targetColMH[j][0]).toUpperCase();
    let currentFormula = lamSach(targetColTonKhoFormulas[j][0]);
    let currentTonKhoValue = lamSach(targetColTonKhoExistingValues[j][0]);
    if (
      sourceSkuSet.has(targetSku) &&
      currentFormula === "" &&
      currentTonKhoValue === ""
    ) {
      let congThucTruXuat = COT.target.cotsXuatKho
        .map((c) => `-${c.sl}${rowNum}`)
        .join("");
      diaChiCanThemCongThucTonKho.push({
        rowNum,
        formula: `=${COT.target.tongNhap}${rowNum}${congThucTruXuat}`,
      });
    }
  }
  if (diaChiCanThemCongThucTonKho.length > 0) {
    for (let item of diaChiCanThemCongThucTonKho) {
      targetSheet
        .getRange(`${COT.target.tonKho}${item.rowNum}`)
        .setFormula(item.formula);
    }
  }

  let targetColTongNhap = targetSheet
    .getRange(
      COT.target.tongNhap + tbd + ":" + COT.target.tongNhap + targetLastRow,
    )
    .getValues();
  let targetColNhap1 = targetSheet
    .getRange(COT.target.nhap1 + tbd + ":" + COT.target.nhap1 + targetLastRow)
    .getValues();
  let targetColNhap2 = targetSheet
    .getRange(COT.target.nhap2 + tbd + ":" + COT.target.nhap2 + targetLastRow)
    .getValues();
  let targetColNhap3 = targetSheet
    .getRange(COT.target.nhap3 + tbd + ":" + COT.target.nhap3 + targetLastRow)
    .getValues();
  let targetColTonKho = targetSheet
    .getRange(COT.target.tonKho + tbd + ":" + COT.target.tonKho + targetLastRow)
    .getValues();
  let targetColChuaHoanThanh = targetSheet
    .getRange(
      COT.target.chuaHoanThanh +
        tbd +
        ":" +
        COT.target.chuaHoanThanh +
        targetLastRow,
    )
    .getValues();
  let targetColSlDonHang = targetSheet
    .getRange(
      COT.target.slDonHang + tbd + ":" + COT.target.slDonHang + targetLastRow,
    )
    .getValues();

  // Đọc các cột xuất kho theo cấu hình mảng
  let xColsData = COT.target.cotsXuatKho.map((c) => ({
    slData: targetSheet
      .getRange(c.sl + tbd + ":" + c.sl + targetLastRow)
      .getValues(),
    ngayData: targetSheet
      .getRange(c.ngay + tbd + ":" + c.ngay + targetLastRow)
      .getValues(),
  }));

  interface TargetRecord {
    index: number;
    viTri: string;
    tongNhap: number;
    tonKho: number;
    chuaHoanThanh: number;
    slDonHang: number;
    daXuatThucTe: number;
    ngayXuat: string[]; // ['25042026', '', '']
    slXuat: string[]; // ['10', '', '']
  }

  let tuDienKho = new Map<string, TargetRecord[]>();
  let targetDataError = "";
  let soDongTuSuaTongNhap = 0;

  for (let j = 0; j < targetColMH.length; j++) {
    let mh = lamSach(targetColMH[j][0]).toUpperCase();
    if (mh === "") continue;
    // Chỉ xử lý/kiểm tra các SKU đang có trong sheet nguồn của lần chạy hiện tại.
    // Tránh bị chặn bởi các dòng dữ liệu đích không liên quan.
    if (!sourceSkuSet.has(mh)) continue;

    let hd = lamSach(targetColHD[j][0]).toUpperCase();
    let vt = lamSach(targetColViTri[j][0]);
    let tnRaw = targetColTongNhap[j][0];
    let n1Raw = targetColNhap1[j][0];
    let n2Raw = targetColNhap2[j][0];
    let n3Raw = targetColNhap3[j][0];
    let tkRaw = targetColTonKho[j][0];
    let chtRaw = targetColChuaHoanThanh[j][0];
    let slDHRaw = targetColSlDonHang[j][0];

    let n1 = parseOptionalNumber(n1Raw);
    if (n1 === null) {
      targetDataError =
        `LỖI TỪ CHỐI: Ô ${COT.target.nhap1}${j + tbd} ` +
        `không phải số hợp lệ. Giá trị hiện tại: '${String(n1Raw)}'`;
      break;
    }
    let n2 = parseOptionalNumber(n2Raw);
    if (n2 === null) {
      targetDataError =
        `LỖI TỪ CHỐI: Ô ${COT.target.nhap2}${j + tbd} ` +
        `không phải số hợp lệ. Giá trị hiện tại: '${String(n2Raw)}'`;
      break;
    }
    let n3 = parseOptionalNumber(n3Raw);
    if (n3 === null) {
      targetDataError =
        `LỖI TỪ CHỐI: Ô ${COT.target.nhap3}${j + tbd} ` +
        `không phải số hợp lệ. Giá trị hiện tại: '${String(n3Raw)}'`;
      break;
    }

    let tn = parseOptionalNumber(tnRaw);
    if (tn === null) {
      targetDataError =
        `LỖI TỪ CHỐI: Ô ${COT.target.tongNhap}${j + tbd} ` +
        `không phải số hợp lệ. Giá trị hiện tại: '${String(tnRaw)}'`;
      break;
    }
    let tongNhapThuCong = Number((n1 + n2 + n3).toFixed(4));
    if (Math.abs(tn - tongNhapThuCong) > 0.0001) {
      let rowNum = j + tbd;
      targetSheet
        .getRange(`${COT.target.tongNhap}${rowNum}`)
        .setFormula(
          `=${COT.target.nhap1}${rowNum}+${COT.target.nhap2}${rowNum}+${COT.target.nhap3}${rowNum}`,
        );
      tn = tongNhapThuCong;
      soDongTuSuaTongNhap++;
    }
    let tkBang = parseOptionalNumber(tkRaw);
    if (tkBang === null) tkBang = 0;
    let cht = parseOptionalNumber(chtRaw);
    if (cht === null) {
      targetDataError =
        `LỖI TỪ CHỐI: Ô ${COT.target.chuaHoanThanh}${j + tbd} ` +
        `không phải số hợp lệ. Giá trị hiện tại: '${String(chtRaw)}'`;
      break;
    }
    let slDH = parseOptionalNumber(slDHRaw);
    if (slDH === null) {
      targetDataError =
        `LỖI TỪ CHỐI: Ô ${COT.target.slDonHang}${j + tbd} ` +
        `không phải số hợp lệ. Giá trị hiện tại: '${String(slDHRaw)}'`;
      break;
    }

    // Thu thập dữ liệu xuất kho của dòng j thành mảng
    let ngayXuat = xColsData.map((c) =>
      chuanHoaNgay(c.ngayData[j][0]).replace(/\D/g, ""),
    );
    let slXuat = xColsData.map((c) => lamSach(c.slData[j][0]));
    let daXuatThucTe = Number(
      slXuat
        .reduce((sum, val) => {
          let n = val === "" ? 0 : Number(val);
          return sum + (isNaN(n) ? 0 : n);
        }, 0)
        .toFixed(4),
    );
    let tonKhoTuTinh = Number((tn - daXuatThucTe).toFixed(4));
    if (tonKhoTuTinh < 0) tonKhoTuTinh = 0;

    let record: TargetRecord = {
      index: j,
      viTri: vt,
      tongNhap: tn,
      tonKho: tonKhoTuTinh,
      chuaHoanThanh: cht,
      slDonHang: slDH,
      daXuatThucTe,
      ngayXuat,
      slXuat,
    };

    let khoa = `${mh}|${hd}`;
    if (!tuDienKho.has(khoa)) {
      tuDienKho.set(khoa, []);
    }
    tuDienKho.get(khoa)!.push(record);
  }
  if (targetDataError !== "") {
    console.log(targetDataError);
    return;
  }
  if (soDongTuSuaTongNhap > 0) {
    console.log(
      `Đã tự sửa công thức cột ${COT.target.tongNhap}: ${soDongTuSuaTongNhap} dòng bị lệch.`,
    );
  }
  console.log(`Đã xây dựng từ điển: ${tuDienKho.size} khóa.`);

  // ==========================================
  // 5. XỬ LÝ LOGIC XUẤT KHO
  // ==========================================
  let ketQua: string[][] = [];
  let ketQuaColors: string[] = [];
  let targetUpdates: {
    row: number;
    col: string;
    val: string | number;
    isDate?: boolean;
  }[] = [];

  let report = {
    total: 0,
    success: 0,
    warningDupDay: 0,
    errorNotFound: 0,
    errorTonKho: 0,
    errorFullCol: 0,
  };

  for (let i = 0; i < mahang.length; i++) {
    let ghiKetQuaDong = (noiDung: string, mau: string) => {
      ketQua.push([noiDung]);
      ketQuaColors.push(mau);
      selectedSheet.getRange(`${COT.source.ketQua}${i + bd}`).setValue(noiDung);
    };
    let currentMaHang = lamSach(mahang[i][0]).toUpperCase();
    if (currentMaHang === "") {
      ghiKetQuaDong("", "");
      continue;
    }

    report.total++;
    let currentHopDong = lamSach(sohopdong[i][0]).toUpperCase();
    let currentMaHangUpper = currentMaHang.toUpperCase();
    let tonKhoToiThieuTheoDong = getTonKhoToiThieuTheoHopDong(currentHopDong);
    let slXuatStr = lamSach(sldonggoi[i][0]);
    let slXuat = slXuatStr === "" ? 0 : Number(slXuatStr);
    let slDaXuatNguonStr = lamSach(slDaXuatNguon[i][0]);
    let slDaXuatNguonSo =
      slDaXuatNguonStr === "" ? null : Number(slDaXuatNguonStr);
    let currentViTri = lamSach(ghiChuSource[i][0]);
    let khoa = `${currentMaHang}|${currentHopDong}`;
    let thongTinCapAB = tachSKUCapAB(currentMaHangUpper);

    if (!tuDienKho.has(khoa)) {
      ghiKetQuaDong(
        "Lỗi: Không tìm thấy mã hàng và hợp đồng trong sheet đích",
        "red",
      );
      report.errorNotFound++;
      continue;
    }

    let danhSach = tuDienKho.get(khoa)!;
    let targetRecord: TargetRecord;
    let chonTheoSlDaXuat = (ds: TargetRecord[]): TargetRecord | null => {
      if (ds.length === 0) return null;
      let dsCoTonKho = ds.filter(
        (d) => d.tongNhap > 0 && !isNaN(d.tonKho) && d.tonKho >= slXuat,
      );
      let tapUngVien = dsCoTonKho.length > 0 ? dsCoTonKho : ds;
      if (slDaXuatNguonSo == null || isNaN(slDaXuatNguonSo)) {
        let coSlot = tapUngVien.filter((d) => d.slXuat.some((x) => x === ""));
        return coSlot.length > 0 ? coSlot[0] : tapUngVien[0];
      }
      let sapXep = tapUngVien
        .map((d) => ({
          record: d,
          lech: Math.abs(Number((d.daXuatThucTe - slDaXuatNguonSo).toFixed(4))),
          coSlot: d.slXuat.some((x) => x === ""),
        }))
        .sort((a, b) => {
          if (a.lech !== b.lech) return a.lech - b.lech;
          if (a.coSlot !== b.coSlot) return a.coSlot ? -1 : 1;
          return b.record.tonKho - a.record.tonKho;
        });
      return sapXep.length > 0 ? sapXep[0].record : null;
    };

    // --- BƯỚC 5.1: TÌM DÒNG ĐÍCH PHÙ HỢP ---
    if (danhSach.length === 1) {
      targetRecord = danhSach[0];
    } else {
      if (currentViTri === "") {
        // Không có vị trí ghi chú: dùng slDaXuat để định vị dòng đích gần đúng nhất.
        let chon = chonTheoSlDaXuat(danhSach);
        if (chon) {
          targetRecord = chon;
        } else {
          ghiKetQuaDong(
            `Lỗi: ${danhSach.length} dòng trùng, không dòng nào đủ tồn kho (${slXuat})`,
            "red",
          );
          report.errorNotFound++;
          continue;
        }
      } else {
        // Phân biệt bằng vị trí ghi chú
        let exactMatches = danhSach.filter((d) => d.viTri === currentViTri);
        if (exactMatches.length === 1) {
          targetRecord = exactMatches[0];
        } else if (exactMatches.length > 1) {
          let chon = chonTheoSlDaXuat(exactMatches);
          if (chon) {
            targetRecord = chon;
          } else {
            ghiKetQuaDong(
              `Lỗi: Có ${exactMatches.length} dòng trùng vị trí '${currentViTri}'`,
              "red",
            );
            report.errorNotFound++;
            continue;
          }
        } else {
          // Khớp một phần
          let partialMatches = danhSach.filter(
            (d) =>
              d.viTri !== "" &&
              currentViTri !== "" &&
              (d.viTri.includes(currentViTri) ||
                currentViTri.includes(d.viTri)),
          );
          if (partialMatches.length === 1) {
            targetRecord = partialMatches[0];
          } else if (partialMatches.length > 1) {
            let chon = chonTheoSlDaXuat(partialMatches);
            if (chon) {
              targetRecord = chon;
            } else {
              ghiKetQuaDong(
                `Lỗi: Vị trí '${currentViTri}' khớp với nhiều dòng`,
                "red",
              );
              report.errorNotFound++;
              continue;
            }
          } else {
            ghiKetQuaDong(
              `Lỗi: Không tìm thấy vị trí '${currentViTri}'`,
              "red",
            );
            report.errorNotFound++;
            continue;
          }
        }
      }
    }

    // --- BƯỚC 5.2: KIỂM TRA ĐIỀU KIỆN XUẤT ---
    if (thongTinCapAB.laSKUCapAB) {
      let thiTruong = lamSach(currentHopDong).toUpperCase().slice(0, 2);
      let khoaCapAB = `${thiTruong}|${currentHopDong}|${thongTinCapAB.skuGoc}`;
      let thongKe = thongKeCapABNguon.get(khoaCapAB) || { tongA: 0, tongB: 0 };
      if (thongKe.tongA <= 0 || thongKe.tongB <= 0) {
        ghiKetQuaDong(
          `Lỗi: SKU cặp A/B chưa đủ đôi trong phiếu (HĐ ${currentHopDong}, SKU gốc ${thongTinCapAB.skuGoc}, A=${thongKe.tongA}, B=${thongKe.tongB})`,
          "red",
        );
        report.errorTonKho++;
        continue;
      }
    }

    let nguongXuatTheoCap = getNguongXuatTheoCap(currentHopDong);
    if (nguongXuatTheoCap > 0 && slXuat > 0) {
      let thiTruong = lamSach(currentHopDong).toUpperCase().slice(0, 2);
      let khoaNhom = `${thiTruong}|${currentHopDong}|${currentMaHang}`;
      let tongNhom = Number(
        (tongXuatTheoNhomNguon.get(khoaNhom) || 0).toFixed(4),
      );
      if (tongNhom < nguongXuatTheoCap) {
        ghiKetQuaDong(
          `Lỗi: Nhóm ${currentHopDong} + ${currentMaHang} của thị trường ${thiTruong} mới xuất ${tongNhom}, chưa đủ ngưỡng xuất theo cặp ${nguongXuatTheoCap}`,
          "red",
        );
        report.errorTonKho++;
        continue;
      }
    }

    if (targetRecord.ngayXuat.includes(dayNorm)) {
      ghiKetQuaDong(
        "Cảnh báo: Mã hàng này đã được xuất kho trước đó (cùng ngày)",
        "yellow",
      );
      report.warningDupDay++;
      continue;
    }

    if (targetRecord.tongNhap === 0) {
      ghiKetQuaDong(
        `Lỗi: Tổng nhập kho bằng 0, không có hàng thực tế để xuất`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    if (targetRecord.tongNhap > targetRecord.slDonHang) {
      ghiKetQuaDong(
        `Lỗi dữ liệu: Tổng nhập (${targetRecord.tongNhap}) vượt số lượng đơn hàng (${targetRecord.slDonHang})`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    if (slXuat > targetRecord.tonKho) {
      ghiKetQuaDong(
        `Lỗi: Số lượng xuất (${slXuat}) lớn hơn tồn kho vật lý hiện hành (${targetRecord.tonKho})`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    // Kiểm tra an toàn: Không xuất vượt số lượng đơn hàng
    let daXuatCacLanTruoc = targetRecord.slXuat.reduce((sum, val) => {
      let num = val === "" ? 0 : Number(val);
      return sum + (isNaN(num) ? 0 : num);
    }, 0);

    // Làm tròn để tránh sai số thập phân của Javascript (VD: 0.1 + 0.2 = 0.30000000000000004)
    daXuatCacLanTruoc = Number(daXuatCacLanTruoc.toFixed(4));
    if (daXuatCacLanTruoc > targetRecord.slDonHang) {
      ghiKetQuaDong(
        `Lỗi dữ liệu: Đã xuất (${daXuatCacLanTruoc}) vượt số lượng đơn hàng (${targetRecord.slDonHang})`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }
    let conLaiCuaDon = Number(
      (targetRecord.slDonHang - daXuatCacLanTruoc).toFixed(4),
    );
    if (conLaiCuaDon <= 0) {
      ghiKetQuaDong(
        `Lỗi: Đơn hàng đã xuất đủ (${daXuatCacLanTruoc}/${targetRecord.slDonHang}), không thể xuất thêm`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    // Tính "mức xuất tối đa cho phép" theo cùng logic soạn hàng ở Auto_Mitsuha:
    // 1) Không vượt tồn kho thực tế
    // 2) Không vượt phần còn lại của đơn hàng
    // 3) Tránh kẹt kho khi phần chưa hoàn thành đang ở mức lẻ < ngưỡng tối thiểu
    let slXuatToiDa = Math.min(targetRecord.tonKho, conLaiCuaDon);
    // Đồng bộ cách tính với Auto_Mitsuha/Auto_Doraemon:
    // - CHT thực tế = max(0, SL đơn hàng - Tổng nhập)
    // - CHT bảng âm/không hợp lệ => fallback về CHT thực tế
    // - Lấy giá trị lớn hơn để tránh đánh giá thiếu phần đang dở dang
    let chtThucTe = targetRecord.slDonHang - targetRecord.tongNhap;
    if (chtThucTe < 0) chtThucTe = 0;
    let chtBang = Number(targetRecord.chuaHoanThanh);
    if (!isFinite(chtBang) || chtBang < 0) chtBang = chtThucTe;
    let chuaHoanThanh = Math.max(chtBang, chtThucTe);
    if (
      chuaHoanThanh > 0 &&
      chuaHoanThanh < tonKhoToiThieuTheoDong &&
      slXuatToiDa > 0
    ) {
      let mucTranChongKetKho = Number(
        (targetRecord.tonKho + chuaHoanThanh - tonKhoToiThieuTheoDong).toFixed(
          4,
        ),
      );
      slXuatToiDa = Math.min(slXuatToiDa, mucTranChongKetKho);
    }
    slXuatToiDa = Number(slXuatToiDa.toFixed(4));

    if (slXuatToiDa <= 0) {
      ghiKetQuaDong(
        `Lỗi: Chưa thể xuất để tránh kẹt kho. Tồn kho=${targetRecord.tonKho}, chưa hoàn thành=${chuaHoanThanh}, ngưỡng tối thiểu=${tonKhoToiThieuTheoDong}`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    if (slXuat > slXuatToiDa) {
      ghiKetQuaDong(
        `Lỗi: Số lượng xuất (${slXuat}) vượt mức cho phép (${slXuatToiDa}) theo quy tắc soạn hàng chống kẹt kho`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    // Kiểm tra số lượng còn lại của đơn sau khi xuất.
    // Nếu xuất xong mà đơn hàng còn dư từ 1 đến 7 cái (nhỏ hơn tồn kho tối thiểu),
    // thì phải cấm xuất để tránh "kẹt kho" ở lần xuất cuối cùng.
    let tongHangConLaiSauXuat = targetRecord.tonKho - slXuat + chuaHoanThanh;
    tongHangConLaiSauXuat = Number(tongHangConLaiSauXuat.toFixed(4));

    if (
      tongHangConLaiSauXuat > 0 &&
      tongHangConLaiSauXuat < tonKhoToiThieuTheoDong
    ) {
      ghiKetQuaDong(
        `Lỗi: Nếu xuất ${slXuat}, đơn hàng sẽ còn dư lại ${tongHangConLaiSauXuat} cái (nhỏ hơn mức tối thiểu ${tonKhoToiThieuTheoDong}). Vui lòng giảm số lượng xuất để chừa lại đủ ${tonKhoToiThieuTheoDong} cái cho lần sau.`,
        "red",
      );
      report.errorTonKho++;
      continue;
    }

    // --- BƯỚC 5.3: TÌM SLOT TRỐNG VÀ GÁN ---
    let slotIndex = targetRecord.slXuat.findIndex((x) => x === "");
    if (slotIndex !== -1) {
      let rowNum = targetRecord.index + tbd;
      let targetColConfig = COT.target.cotsXuatKho[slotIndex];

      targetUpdates.push({ row: rowNum, col: targetColConfig.sl, val: slXuat });
      targetUpdates.push({
        row: rowNum,
        col: targetColConfig.ngay,
        val: dayToUse,
        isDate: true,
      });

      // Cập nhật record tạm thời trên RAM để tránh trùng lặp nội bộ trong 1 lần chạy
      targetRecord.slXuat[slotIndex] = String(slXuat);
      targetRecord.ngayXuat[slotIndex] = dayNorm;

      // Trừ lùi tồn kho và cập nhật số lượng chưa hoàn thành trên RAM
      // để các dòng xuất tiếp theo (nếu trỏ cùng vào dòng đích này) kiểm tra đúng số dư
      targetRecord.tonKho = Number((targetRecord.tonKho - slXuat).toFixed(4));

      // Xóa phần trừ lùi chuaHoanThanh vì "Chưa hoàn thành" là hàng trên chuyền sản xuất.
      // Xuất hàng thành phẩm không làm thay đổi số lượng đang sản xuất.

      ghiKetQuaDong(`Xuất kho ${slotIndex + 1} thành công`, "green");
      report.success++;
    } else {
      ghiKetQuaDong(
        `Lỗi: Cả ${COT.target.cotsXuatKho.length} cột xuất kho đều đã có giá trị`,
        "red",
      );
      report.errorFullCol++;
    }
  }

  // ==========================================
  // 6. GHI KẾT QUẢ VÀ ĐỊNH DẠNG TỰ ĐỘNG
  // ==========================================
  // Ghi theo khối (cùng cột, dòng liên tiếp) — giảm số lần setNumberFormatLocal/setValue, tránh timeout script
  ghiTargetUpdatesTheoLo(targetSheet, targetUpdates);

  ketQuaRange.setValues(ketQua);
  ketQuaRange.clearAllConditionalFormats();

  // Helper hàm bôi màu CF
  let addCF = (text: string, bgColor: string, fontColor: string) => {
    let cf = ketQuaRange.addConditionalFormat(
      ExcelScript.ConditionalFormatType.containsText,
    );
    cf.getTextComparison().getFormat().getFill().setColor(bgColor);
    cf.getTextComparison().getFormat().getFont().setColor(fontColor);
    cf.getTextComparison().setRule({
      operator: ExcelScript.ConditionalTextOperator.contains,
      text: text,
    });
  };

  addCF("Thành công", "#00B050", "#FFFFFF"); // Xanh lá
  addCF("Lỗi", "#FF0000", "#FFFFFF"); // Đỏ
  addCF("Cảnh báo", "#FFC000", "#000000"); // Vàng cam

  console.log(
    `\n=========================================\n` +
      `         BÁO CÁO KẾT QUẢ XUẤT KHO\n` +
      `=========================================\n` +
      `• Tổng số dòng xử lý   : ${report.total}\n` +
      `• Thành công (Xanh)    : ${report.success}\n` +
      `• Đã xuất trong ngày   : ${report.warningDupDay}\n` +
      `• Lỗi thiếu tồn kho    : ${report.errorTonKho}\n` +
      `• Không tìm thấy ở đích: ${report.errorNotFound}\n` +
      `• Bị đầy cột xuất kho  : ${report.errorFullCol}\n` +
      `=========================================`,
  );
}
