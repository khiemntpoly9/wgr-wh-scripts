import { beforeAll, describe, expect, it } from "vitest";

let T;
beforeAll(async () => {
  await import("../finished-goods/Auto_Sanji_DoiSoat.ts");
  T = globalThis.__AUTO_SANJI_DOI_SOAT_TESTABLES__;
});

const apiHeader = [
  "STT", "Mã đơn hàng", "Thị trường", "Mã phân loại", "Mã hợp đồng", "SKU", "Tên hàng", "Vị trí",
  "Số lượng đơn hàng", "Tổng nhập", "Tổng xuất", "Tồn", "Trọng lượng", "Tổng trọng lượng",
  "Dài", "Rộng", "Cao", "Thể tích", "Tổng thể tích",
];
const apiRow = (contract, sku, locations, stock, name = "Bàn", orderQty = 10) => [
  1, "PO-1", contract.slice(0, 2), "F", contract, sku, name, locations,
  orderQty, stock, 0, stock, "", "", "", "", "", "", "",
];
const apiBody = (rows) => ({
  format: "auto_sanji_backend_v1",
  sheet: {
    writes: [{
      address: "A1:S20",
      values: [
        ["THỐNG KÊ TỒN TỪ HỆ THỐNG"],
        [""],
        ["TỔNG CỘNG"],
        ["Thị trường", "Số SKU tồn", "Tổng tồn"],
        ["EU", rows.length, 0],
        [""],
        ["", "", "", "", "TỔNG CỘNG"],
        apiHeader,
        ...rows,
      ],
    }],
  },
});
const col = (values) => values.map((v) => [v]);
const excelSource = (rows) => ({
  maHang: col(rows.map((r) => r.sku)),
  tenHang: col(rows.map((r) => r.name ?? "")),
  slDonHang: col(rows.map((r) => r.order ?? 10)),
  hopDong: col(rows.map((r) => r.contract)),
  viTri: col(rows.map((r) => r.location ?? "")),
  nhapKho: [col(rows.map((r) => r.in ?? 0)), col(rows.map(() => ""))],
  xuatKho: [col(rows.map((r) => r.out ?? 0))],
});

describe("Auto_Sanji_DoiSoat location parsing", () => {
  it("normalizes location codes like the backend parser", () => {
    expect(T.doiSoatTachViTri("B31B30A7")).toEqual(["A7", "B30", "B31"]);
    expect(T.doiSoatTachViTri("d-16, A07")).toEqual(["A7", "D16"]);
    expect(T.doiSoatTachViTri("")).toEqual([]);
    expect(T.doiSoatTachViTri("Kho tam")).toEqual(["KHOTAM"]);
    expect(T.doiSoatTachViTriApi("KHOTAM A7 D-16")).toEqual(["A7", "D16", "KHOTAM"]);
    expect(T.doiSoatLocKhongDefault(["DEFAULT", "A7", "default"])).toEqual(["A7"]);
    expect(T.doiSoatTaoPayload({ marketCodes: ["eu"], apiUrl: "", apiToken: "" })).toMatchObject({
      market_codes: ["EU"],
      use_effective_locations: true,
    });
  });
});

describe("Auto_Sanji_DoiSoat comparison", () => {
  it("reads the detail table from the Auto_Sanji_Super plan", () => {
    const map = T.doiSoatDocDuLieuApi(apiBody([apiRow("EUF01", "SKU-1", "A7 B30", 5)]));
    expect(map.size).toBe(1);
    expect(map.get("EUF01|SKU-1")).toMatchObject({
      thiTruong: "EU", maDonHang: "PO-1", slDonHang: 10, ton: 5, viTri: ["A7", "B30"], viTriHienThi: "A7 B30",
    });
    expect(T.doiSoatDocDuLieuApi({ format: "auto_sanji_backend_v1", sheet: null }).size).toBe(0);
    expect(T.doiSoatDocDuLieuApi({ format: "auto_sanji_backend_v1", sheet: { writes: [] } })).toBeNull();
  });

  it("aggregates Excel rows by contract + SKU without summing order quantity and filters markets", () => {
    const map = T.doiSoatDocDuLieuExcel(
      excelSource([
        { sku: "SKU-1", contract: "EUF01", in: 10, out: 3, location: "A7" },
        { sku: " sku-1 ", contract: "euf01", in: "2", location: "B30" },
        { sku: "SKU-2", contract: "JPT01", in: 4 },
        { sku: "SKU-3", contract: "", in: 4 },
        { sku: "SKU-4", contract: "EUF02", in: 3, order: 0, location: "C1" },
        { sku: "SKU-4", contract: "EUF02", in: 2, order: 15, location: "C2" },
      ]),
      ["eu"],
    );
    expect(Array.from(map.keys())).toEqual(["EUF01|SKU-1", "EUF02|SKU-4"]);
    expect(map.get("EUF01|SKU-1")).toMatchObject({ ton: 9, slDonHang: 10, viTri: ["A7", "B30"], viTriHienThi: "A7 B30" });
    expect(map.get("EUF02|SKU-4")).toMatchObject({ ton: 5, slDonHang: 15, viTri: ["C1", "C2"], viTriHienThi: "C1 C2" });
  });

  it("classifies stock and location discrepancies and sorts mismatches first", () => {
    const heThong = T.doiSoatDocDuLieuApi(apiBody([
      apiRow("EUF01", "OK", "A7", 5),
      apiRow("EUF01", "STOCK", "A7", 5),
      apiRow("EUF01", "LOC", "A7 B30", 5),
      apiRow("EUF01", "BOTH", "A7", 5),
      apiRow("EUF01", "ONLYHT", "A7", 2),
    ]));
    const excel = T.doiSoatDocDuLieuExcel(excelSource([
      { sku: "OK", contract: "EUF01", in: 5, location: "a-07" },
      { sku: "STOCK", contract: "EUF01", in: 8, out: 1, location: "A7" },
      { sku: "LOC", contract: "EUF01", in: 5, location: "B30 C1" },
      { sku: "BOTH", contract: "EUF01", in: 3, location: "C2" },
      { sku: "ONLYEX", contract: "EUF01", in: 4, location: "A1" },
      { sku: "ZERO", contract: "EUF01", in: 4, out: 4 },
    ]), []);

    const result = T.doiSoatSoSanh(heThong, excel, true);
    expect(result.map((r) => [r.sku, r.trangThai])).toEqual([
      ["BOTH", "Lệch tồn + vị trí"],
      ["STOCK", "Lệch tồn"],
      ["LOC", "Lệch vị trí"],
      ["ONLYHT", "Chỉ có trên hệ thống"],
      ["ONLYEX", "Chỉ có trên Excel"],
      ["OK", "Khớp"],
    ]);
    const loc = result.find((r) => r.sku === "LOC");
    expect(loc).toMatchObject({ viTriThieuTrenExcel: "A7", viTriThuaTrenExcel: "C1", chenhLech: 0 });
    expect(result.find((r) => r.sku === "STOCK").chenhLech).toBe(2);
    expect(result.find((r) => r.sku === "ONLYHT")).toMatchObject({ tonExcel: 0, chenhLech: -2 });

    expect(T.doiSoatSoSanh(heThong, excel, false).find((r) => r.sku === "LOC").trangThai).toBe("Khớp");
    expect(T.doiSoatDemTrangThai(result)["Khớp"]).toBe(1);
    expect(T.doiSoatGomKhoi(result).length).toBe(6);
    expect(T.doiSoatDongChiTiet(result)[0]).toHaveLength(18);
  });

  it("flags order quantity discrepancies only when both sources have the SKU", () => {
    const heThong = T.doiSoatDocDuLieuApi(apiBody([
      apiRow("EUF01", "QTY", "A7", 5, "Bàn", 12),
      apiRow("EUF01", "QTYSTOCK", "A7", 5, "Bàn", 12),
      apiRow("EUF01", "SAME", "A7", 5, "Bàn", 12),
      apiRow("EUF01", "ONLYHT", "A7", 5, "Bàn", 12),
    ]));
    const excel = T.doiSoatDocDuLieuExcel(excelSource([
      { sku: "QTY", contract: "EUF01", in: 5, location: "A7", order: 10 },
      { sku: "QTYSTOCK", contract: "EUF01", in: 6, location: "A7", order: "15" },
      { sku: "SAME", contract: "EUF01", in: 5, location: "A7", order: 12 },
      { sku: "ONLYEX", contract: "EUF01", in: 5, location: "A7", order: 3 },
    ]), []);

    const result = T.doiSoatSoSanh(heThong, excel, true, true);
    const bySku = (sku) => result.find((r) => r.sku === sku);
    expect(bySku("QTY")).toMatchObject({
      trangThai: "Lệch SL đơn hàng", lechSlDon: true, slDonHeThong: 12, slDonExcel: 10, chenhLechSlDon: -2,
    });
    expect(bySku("QTYSTOCK")).toMatchObject({ trangThai: "Lệch tồn", lechSlDon: true, chenhLechSlDon: 3 });
    expect(bySku("SAME")).toMatchObject({ trangThai: "Khớp", lechSlDon: false });
    expect(bySku("ONLYHT").lechSlDon).toBe(false);
    expect(bySku("ONLYEX").lechSlDon).toBe(false);

    const row = T.doiSoatDongChiTiet([bySku("QTYSTOCK")])[0];
    expect(row.slice(10, 13)).toEqual([12, 15, 3]);
    expect(row[17]).toBe("Lệch SL đơn hàng");
    expect(T.doiSoatDongChiTiet([bySku("QTY")])[0][17]).toBe("");

    expect(T.doiSoatSoSanh(heThong, excel, true, false).find((r) => r.sku === "QTY").trangThai).toBe("Khớp");
  });
});
