import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_MITSUHA_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_Mitsuha test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Auto_Mitsuha.ts");
});

const toExcelSerial = (day, month, year) => {
  let jsDate = new Date(Date.UTC(year, month - 1, day));
  return Math.floor(jsDate.getTime() / 86400000) + 25569;
};

describe("Auto_Mitsuha helpers", () => {
  it("cleans values and parses Excel-style numbers", () => {
    let { lamSach, parseSoTuExcel } = getTestables();

    expect(lamSach("  AB \u00A0\u200B 12 ")).toBe("AB12");
    expect(parseSoTuExcel("1.234,56")).toBe(1234.56);
    expect(parseSoTuExcel("1,234.56")).toBe(1234.56);
    expect(parseSoTuExcel("1234")).toBe(1234);
    expect(parseSoTuExcel("abc")).toBe(0);
  });

  it("normalizes Excel serial dates and leaves strings trimmed", () => {
    let { chuanHoaNgay } = getTestables();

    expect(chuanHoaNgay(toExcelSerial(12, 5, 2026))).toBe("12/05/2026");
    expect(chuanHoaNgay(" 08/05/2026 ")).toBe("08/05/2026");
  });

  it("picks the latest inbound date among receipt columns", () => {
    let { layNgayNhapMoiNhatTuCacCot, ngayKeySangExcelSerial } = getTestables();

    expect(
      layNgayNhapMoiNhatTuCacCot([
        toExcelSerial(1, 3, 2026),
        "15/05/2026",
        toExcelSerial(10, 4, 2026),
      ]),
    ).toBe(toExcelSerial(15, 5, 2026));
    expect(layNgayNhapMoiNhatTuCacCot(["", null, ""])).toBe("");
    expect(ngayKeySangExcelSerial("2026-05-15")).toBe(toExcelSerial(15, 5, 2026));
    expect(ngayKeySangExcelSerial("")).toBe("");
  });

  it("builds pending keys using contract, sku, order quantity, and normalized location", () => {
    let { chuanHoaViTri, taoKhoaPendingTheoNhom, taoKhoaPendingTheoViTri } =
      getTestables();

    expect(chuanHoaViTri(" b044-c72 ")).toBe("B44 C72");
    expect(chuanHoaViTri("B25 A15")).toBe("A15 B25");
    expect(chuanHoaViTri("B26 A15")).toBe("A15 B26");

    let khoaNhom = taoKhoaPendingTheoNhom(" ue-01 ", " 901019A ", 16, 8);
    expect(khoaNhom).toBe("UE|ue-01|901019A|16");
    expect(taoKhoaPendingTheoViTri(khoaNhom, " b044-c72 ")).toBe(
      "UE|ue-01|901019A|16|B44 C72",
    );
  });

  it("marks dropped goods as Pre-packing-list minus Xuất công", () => {
    let {
      taoKhoaHangRot,
      chiSoCotExcelHangRot,
      taoMetaHangRotAR,
      timDongTieuDeHangRotAR,
      congDonSoLuongTheoKhoaHangRot,
      tinhSlHangRotTheoNhom,
      ganNhanCotHangRot,
    } = getTestables();

    expect(taoKhoaHangRot(" eu-01 ", " 901019A ")).toBe("EU-01|901019A");
    expect(chiSoCotExcelHangRot("D")).toBe(3);
    expect(chiSoCotExcelHangRot("K")).toBe(10);
    expect(chiSoCotExcelHangRot("R")).toBe(17);

    let header = new Array(18).fill("");
    header[3] = "Mã hợp đồng";
    header[4] = "SKU";
    header[9] = "Đã xuất đi";
    header[10] = "Số lượng soạn";
    header[17] = "Vị trí";

    let dongAR = (hd, sku, soan, daXuat) => {
      let row = new Array(18).fill("");
      row[3] = hd;
      row[4] = sku;
      row[9] = daXuat;
      row[10] = soan;
      return row;
    };

    let preRows = [
      header,
      dongAR("EU-01", "901019A", 16, 0),
      dongAR("EU-01", "901019A", 4, 0),
      dongAR("EU-02", "901020B", 10, 0),
    ];
    let xuatRows = [
      header,
      dongAR("EU-01", "901019A", 16, 12),
      dongAR("EU-02", "901020B", 10, 10),
    ];

    expect(timDongTieuDeHangRotAR(preRows)).toBe(0);
    let preMeta = taoMetaHangRotAR(0, false);
    let xuatMeta = taoMetaHangRotAR(0, true);
    expect(preMeta.cotHopDong).toBe(3);
    expect(preMeta.cotMaHang).toBe(4);
    expect(preMeta.cotSoLuongSoan).toBe(10);

    let soanTruoc = congDonSoLuongTheoKhoaHangRot(preRows, preMeta);
    let daXuat = congDonSoLuongTheoKhoaHangRot(xuatRows, xuatMeta);
    expect(soanTruoc["EU-01|901019A"]).toBe(20);
    expect(daXuat["EU-01|901019A"]).toBe(12);
    let hangRot = tinhSlHangRotTheoNhom(soanTruoc, daXuat);
    expect(hangRot["EU-01|901019A"]).toBe(8);
    expect(hangRot["EU-02|901020B"]).toBeUndefined();

    expect(ganNhanCotHangRot("RỚT", 8, 8)).toBe("RỚT");
    expect(ganNhanCotHangRot("RỚT", 8, 16)).toBe("RỚT 8/16");
  });

  it("loads pending stock hold from Pre-packing-list A-R columns", () => {
    let { napPendingTuBangAR } = getTestables();

    let header = new Array(18).fill("");
    header[3] = "Mã hợp đồng";
    header[4] = "SKU";
    header[7] = "Số lượng đơn hàng";
    header[10] = "Số lượng soạn";
    header[17] = "Vị trí";

    let row = (hd, sku, slDon, soan, viTri) => {
      let r = new Array(18).fill("");
      r[3] = hd;
      r[4] = sku;
      r[7] = slDon;
      r[10] = soan;
      r[17] = viTri;
      return r;
    };

    let nap = napPendingTuBangAR(
      [
        header,
        row("EU-01", "901019A", 16, 10, "A15"),
        row("EU-01", "901019A", 16, 6, "B25"),
      ],
      0,
    );

    expect(nap.soDongPendingHopLe).toBe(2);
    expect(nap.tongPendingTrongSheet).toBe(16);
    expect(nap.pendingTheoNhom["EU|EU-01|901019A|16"]).toBe(16);
    expect(nap.pendingTheoViTri["EU|EU-01|901019A|16|A15"]).toBe(10);
    expect(nap.pendingTheoViTri["EU|EU-01|901019A|16|B25"]).toBe(6);
  });

  it("does not hold dropped repeats as pending so the sku can be picked again", () => {
    let { phanTichPendingVaHangRotTuBangAR } = getTestables();

    let header = new Array(18).fill("");
    header[3] = "Mã hợp đồng";
    header[4] = "SKU";
    header[7] = "Số lượng đơn hàng";
    header[10] = "Số lượng soạn";

    let row = (ngay, phieu, hd, sku, slDon, soan) => {
      let r = new Array(18).fill("");
      r[0] = ngay;
      r[1] = phieu;
      r[3] = hd;
      r[4] = sku;
      r[7] = slDon;
      r[10] = soan;
      return r;
    };

    let preRows = [
      header,
      row("12/08/2026", 429, "EU-01", "900872", 15, 15),
      row("12/08/2026", 430, "EU-01", "900872", 15, 15),
      row("13/08/2026", 431, "EU-01", "900872", 15, 15),
      row("14/08/2026", 434, "EU-01", "901019A", 8, 8),
      row("14/08/2026", 435, "EU-01", "901019A", 8, 4),
    ];

    let ketQua = phanTichPendingVaHangRotTuBangAR(preRows, 0, [header], 0);
    expect(ketQua.pendingTheoNhom["EU|EU-01|900872|15"]).toBeUndefined();
    expect(ketQua.slHangRotTheoNhom["EU-01|900872"]).toBe(15);
    expect(ketQua.thongTinHangRotTheoNhom["EU-01|900872"]).toEqual({
      soPhieu: "431",
      ngayKey: "2026-08-13",
    });
    expect(ketQua.pendingTheoNhom["EU|EU-01|901019A|8"]).toBe(12);
  });

  it("does not double-count pending already reflected on Thành phẩm exports", () => {
    let { dieuChinhTonVaDaXuatSauPending } = getTestables();

    // Pre 50, Xuất công 20, Thành phẩm đã ghi 20 → còn 30 cần soạn, tồn 30 khả dụng.
    expect(
      dieuChinhTonVaDaXuatSauPending(20, 10, 20, 20),
    ).toEqual({
      daXuatTamTinh: 20,
      tonKhoKhaDung: 30,
    });

    // Thành phẩm chưa cập nhật: pending nhóm (EU-1) vẫn chiếm phần đơn và tồn.
    expect(
      dieuChinhTonVaDaXuatSauPending(0, 30, 20, 20),
    ).toEqual({
      daXuatTamTinh: 20,
      tonKhoKhaDung: 30,
    });

    // Pending nhóm lớn hơn phần đã ghi trên Thành phẩm.
    expect(
      dieuChinhTonVaDaXuatSauPending(32, 0, 50, 50),
    ).toEqual({
      daXuatTamTinh: 50,
      tonKhoKhaDung: 32,
    });

    // EU-1 đã soạn 20, EU-2 cùng SKU: pending nhóm 20, Thành phẩm 0 → H hiển thị 20.
    expect(
      dieuChinhTonVaDaXuatSauPending(0, 40, 20, 20),
    ).toEqual({
      daXuatTamTinh: 20,
      tonKhoKhaDung: 40,
    });

    // Thành phẩm đã ghi nhiều hơn pending tại dòng — dùng max(TP, pending nhóm).
    expect(
      dieuChinhTonVaDaXuatSauPending(30, 10, 20, 20),
    ).toEqual({
      daXuatTamTinh: 30,
      tonKhoKhaDung: 30,
    });
  });

  it("holds exported Xuất công qty as pending so loaded goods are not packed again", () => {
    let { phanTichPendingVaHangRotTuBangAR } = getTestables();

    let header = new Array(18).fill("");
    header[3] = "Mã hợp đồng";
    header[4] = "SKU";
    header[7] = "Số lượng đơn hàng";
    header[10] = "Số lượng soạn";

    let row = (ngay, phieu, hd, sku, slDon, soan, daXuat) => {
      let r = new Array(18).fill("");
      r[0] = ngay;
      r[1] = phieu;
      r[3] = hd;
      r[4] = sku;
      r[7] = slDon;
      r[9] = daXuat || 0;
      r[10] = soan;
      return r;
    };

    let preRows = [
      header,
      row("14/08/2026", 434, "EU-01", "900872", 15, 15, 0),
    ];
    let xuatRows = [
      header,
      row("15/08/2026", 434, "EU-01", "900872", 15, 15, 15),
    ];

    let du = phanTichPendingVaHangRotTuBangAR(preRows, 0, xuatRows, 0);
    expect(du.pendingTheoNhom["EU|EU-01|900872|15"]).toBe(15);
    expect(du.slHangRotTheoNhom["EU-01|900872"]).toBeUndefined();

    let thieu = phanTichPendingVaHangRotTuBangAR(
      [header, row("14/08/2026", 434, "EU-01", "900872", 15, 15, 0)],
      0,
      [header, row("15/08/2026", 434, "EU-01", "900872", 10, 15, 0)],
      0,
    );
    expect(thieu.pendingTheoNhom["EU|EU-01|900872|15"]).toBe(10);
    expect(thieu.slHangRotTheoNhom["EU-01|900872"]).toBe(5);
  });

  it("uses Xuất công column H as exported qty when J/K are smaller", () => {
    let { phanTichPendingVaHangRotTuBangAR } = getTestables();

    let header = new Array(18).fill("");
    header[3] = "Mã hợp đồng";
    header[4] = "SKU";
    header[7] = "Số lượng xuất đi";
    header[10] = "Số lượng soạn";

    let row = (ngay, phieu, hd, sku, cotH, soan, daXuat) => {
      let r = new Array(18).fill("");
      r[0] = ngay;
      r[1] = phieu;
      r[3] = hd;
      r[4] = sku;
      r[7] = cotH;
      r[9] = daXuat || 0;
      r[10] = soan;
      return r;
    };

    let ketQua = phanTichPendingVaHangRotTuBangAR(
      [header, row("15/08/2026", 434, "EU-01", "943996", 10, 10, 0)],
      0,
      [header, row("15/08/2026", 434, "EU-01", "943996", 25, 10, 0)],
      0,
    );
    expect(ketQua.pendingTheoNhom["EU|EU-01|943996|10"]).toBe(25);
    expect(ketQua.slHangRotTheoNhom["EU-01|943996"]).toBeUndefined();
  });

  it("does not mark fully loaded phiếu+SKU as dropped (943996 / 434)", () => {
    let { phanTichPendingVaHangRotTuBangAR } = getTestables();

    let header = new Array(18).fill("");
    header[3] = "Mã hợp đồng";
    header[4] = "SKU";
    header[7] = "Số lượng đơn hàng";
    header[10] = "Số lượng soạn";

    let row = (ngay, phieu, hd, sku, slDon, soan, daXuat) => {
      let r = new Array(18).fill("");
      r[0] = ngay;
      r[1] = phieu;
      r[3] = hd;
      r[4] = sku;
      r[7] = slDon;
      r[9] = daXuat || 0;
      r[10] = soan;
      return r;
    };

    let ketQua = phanTichPendingVaHangRotTuBangAR(
      [header, row("15/08/2026", 434, "EU-01", "943996", 12, 12, 0)],
      0,
      [header, row("15/08/2026", "434.0", "EU-01", 943996, 12, 12, 12)],
      0,
    );
    expect(ketQua.pendingTheoNhom["EU|EU-01|943996|12"]).toBe(12);
    expect(ketQua.slHangRotTheoNhom["EU-01|943996"]).toBeUndefined();
    expect(ketQua.thongTinHangRotTheoNhom["EU-01|943996"]).toBeUndefined();

    let lechHd = phanTichPendingVaHangRotTuBangAR(
      [header, row("15/08/2026", 434, "EU55005", "943996A", 12, 12, 0)],
      0,
      [header, row("15/08/2026", 434, "EU55005-1", "943996", 12, 12, 12)],
      0,
    );
    expect(lechHd.slHangRotTheoNhom["EU55005|943996A"]).toBeUndefined();
    expect(lechHd.pendingTheoNhom["EU|EU55005|943996A|12"]).toBe(12);
  });

  it("formats dropped slip and latest pack date", () => {
    let { ganChuoiPhieuNgayHangRot } = getTestables();
    expect(ganChuoiPhieuNgayHangRot("431", "2026-08-13")).toBe(
      "431 — 13/08/2026",
    );
  });

  it("subtracts pending by location first, then by group remainder", () => {
    let { rutSoLuongPending } = getTestables();

    let pendingTheoNhom = {
      "UE|HD01|901019A|16": 7,
    };
    let pendingTheoViTri = {
      "UE|HD01|901019A|16|B44": 3,
    };

    let ketQua = rutSoLuongPending(
      10,
      "UE|HD01|901019A|16",
      "B44",
      pendingTheoNhom,
      pendingTheoViTri,
    );

    expect(ketQua).toEqual({
      tonKhoSauPending: 3,
      soLuongPendingDaTru: 7,
    });
    expect(pendingTheoNhom["UE|HD01|901019A|16"]).toBe(0);
    expect(pendingTheoViTri["UE|HD01|901019A|16|B44"]).toBe(0);
  });

  it("skuKhongDuocSoan — số 6 chữ số loại cả A/B, chuỗi chỉ khớp đúng mã", () => {
    let { taoTapSkuKhongDuocSoan, laSkuKhongDuocSoan } = getTestables();
    let tap = taoTapSkuKhongDuocSoan([924125, "901019A"]);

    expect(laSkuKhongDuocSoan("924125A", tap)).toBe(true);
    expect(laSkuKhongDuocSoan("924125B", tap)).toBe(true);
    expect(laSkuKhongDuocSoan("901019A", tap)).toBe(true);
    expect(laSkuKhongDuocSoan("901019B", tap)).toBe(false);
  });

  it("matches pending even if exported quantity snapshot changed", () => {
    let { taoKhoaPendingTheoNhom, rutSoLuongPending } = getTestables();

    let khoaTheoNguon = taoKhoaPendingTheoNhom("EU-903980", "903980", 50, 0);
    let khoaTheoPending = taoKhoaPendingTheoNhom("EU-903980", "903980", 50, 32);
    expect(khoaTheoPending).toBe(khoaTheoNguon);

    let pendingTheoNhom = {
      [khoaTheoPending]: 50, // pending-auto có 2 dòng 32 + 18
    };
    let pendingTheoViTri = {};

    let lan1 = rutSoLuongPending(
      32,
      khoaTheoNguon,
      "",
      pendingTheoNhom,
      pendingTheoViTri,
    );
    expect(lan1).toEqual({
      tonKhoSauPending: 0,
      soLuongPendingDaTru: 32,
    });
    expect(pendingTheoNhom[khoaTheoNguon]).toBe(18);

    let lan2 = rutSoLuongPending(
      18,
      khoaTheoNguon,
      "",
      pendingTheoNhom,
      pendingTheoViTri,
    );
    expect(lan2).toEqual({
      tonKhoSauPending: 0,
      soLuongPendingDaTru: 18,
    });
    expect(pendingTheoNhom[khoaTheoNguon]).toBe(0);
  });

  it("subtracts pending by group even when pending rows are different locations", () => {
    let { rutSoLuongPending } = getTestables();

    let pendingTheoNhom = {
      "EU|EU-903980|903980|50": 50,
    };
    let pendingTheoViTri = {
      "EU|EU-903980|903980|50|A1": 32,
      "EU|EU-903980|903980|50|B2": 18,
    };

    let ketQua = rutSoLuongPending(
      50,
      "EU|EU-903980|903980|50",
      "C3",
      pendingTheoNhom,
      pendingTheoViTri,
    );

    expect(ketQua).toEqual({
      tonKhoSauPending: 0,
      soLuongPendingDaTru: 50,
    });
    expect(pendingTheoNhom["EU|EU-903980|903980|50"]).toBe(0);
  });

  it("detects groups completed by pending plus current auto wave", () => {
    let { tinhTapNhomHoanTatCoPending } = getTestables();

    let tapNhom = tinhTapNhomHoanTatCoPending(
      [
        {
          khoaNhom: "UE|HD01|929666|16",
          conLaiTruocXuat: 8,
          soLuongXuat: 8,
        },
        {
          khoaNhom: "UE|HD02|901019A|12",
          conLaiTruocXuat: 12,
          soLuongXuat: 10,
        },
      ],
      {
        "UE|HD01|929666|16": 8,
        "UE|HD02|901019A|12": 2,
      },
    );

    expect(tapNhom.has("UE|HD01|929666|16")).toBe(true);
    expect(tapNhom.has("UE|HD02|901019A|12")).toBe(false);
  });

  it("blocks a contract+SKU group when one production split is not fully received", () => {
    let {
      taoKhoaNhomHopDongSKU,
      tinhChuaHoanThanhDong,
      tongHopChuaHoanThanhTheoNhomHopDongSKU,
    } = getTestables();

    expect(taoKhoaNhomHopDongSKU("UE-HD01", "929666")).toBe(
      "UE|UE-HD01|929666",
    );
    expect(tinhChuaHoanThanhDong(8, 8, 0)).toBe(0);
    expect(tinhChuaHoanThanhDong(7, 3, 0)).toBe(4);

    let tongHop = tongHopChuaHoanThanhTheoNhomHopDongSKU([
      {
        hopDong: "UE-HD01",
        maHang: "929666",
        slDonHang: 8,
        tongNhap: 8,
        chuaHoanThanhBang: 0,
      },
      {
        hopDong: "UE-HD01",
        maHang: "929666",
        slDonHang: 7,
        tongNhap: 3,
        chuaHoanThanhBang: 0,
      },
    ]);

    expect(tongHop["UE|UE-HD01|929666"]).toBe(4);
  });

  it("builds pre-picking-list sheet names from market filters", () => {
    let { taoTenSheetPrePickingList } = getTestables();

    expect(taoTenSheetPrePickingList(["UE"])).toBe("pre-picking-list-UE");
    expect(taoTenSheetPrePickingList([])).toBe("pre-picking-list-ALL");
    expect(taoTenSheetPrePickingList(["UE", "EU", "A/B"])).toBe(
      "pre-picking-list-UE-EU-A-B",
    );
    expect(taoTenSheetPrePickingList(["UE".repeat(20)]).length).toBeLessThanOrEqual(
      31,
    );
  });

  it("defaults blank or zero height dimensions to 1.1", () => {
    let { chuanHoaChieuCaoKichThuoc } = getTestables();

    expect(chuanHoaChieuCaoKichThuoc("")).toBe(1.1);
    expect(chuanHoaChieuCaoKichThuoc(0)).toBe(1.1);
    expect(chuanHoaChieuCaoKichThuoc("0")).toBe(1.1);
    expect(chuanHoaChieuCaoKichThuoc("2,5")).toBe(2.5);
  });

  it("matches product specs by fuzzy token overlap on product names", () => {
    let {
      chuanHoaTenSanPham,
      tachTokenTenSanPham,
      tinhDiemKhopTenMo,
      timThongSoTheoTenMo,
    } = getTestables();

    let tenTraCuu = chuanHoaTenSanPham("B0W ac 1p bc din tab rect I");
    let tokensTraCuu = tachTokenTenSanPham(tenTraCuu);

    let tenBc = chuanHoaTenSanPham("B0W ac 1p bc din tab life l");
    let tenCe = chuanHoaTenSanPham("B0W ac 1p ce din tab life l");

    expect(tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenBc))).toBeCloseTo(0.6);
    expect(tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenBc))).toBeGreaterThan(
      tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenCe)),
    );

    let danhSach = [tenBc, tenCe].map((tenChuan, i) => ({
      tenChuan,
      tokens: tachTokenTenSanPham(tenChuan),
      thongSo: { kichThuoc: `k-${i}`, trongLuong: 2, theTich: 0.2 },
      coDuLieu: true,
    }));

    expect(timThongSoTheoTenMo(tenTraCuu, danhSach)?.kichThuoc).toBe("k-0");
  });
});
