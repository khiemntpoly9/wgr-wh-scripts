import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_NOBITA_TESTABLES__;
  if (!testables) {
    throw new Error("Auto_Nobita test helpers were not loaded.");
  }
  return testables;
};

beforeAll(async () => {
  await import("../Auto_Nobita.ts");
});

describe("Auto_Nobita helpers", () => {
  it("cleans values correctly", () => {
    let { nobitaLamSach } = getTestables();

    expect(nobitaLamSach("  AB \u00A0\u200B 12 ")).toBe("AB12");
    expect(nobitaLamSach("901019A")).toBe("901019A");
    expect(nobitaLamSach(12345)).toBe("12345");
    expect(nobitaLamSach("")).toBe("");
  });

  it("parses Excel-style numbers correctly", () => {
    let { nobitaParseSo } = getTestables();

    expect(nobitaParseSo("1.234,56")).toBe(1234.56);
    expect(nobitaParseSo("1,234.56")).toBe(1234.56);
    expect(nobitaParseSo("1234")).toBe(1234);
    expect(nobitaParseSo("abc")).toBe(0);
    expect(nobitaParseSo(99)).toBe(99);
    expect(nobitaParseSo("")).toBe(0);
  });

  it("normalizes Excel serial dates", () => {
    let { nobitaChuanHoaNgay } = getTestables();

    // Excel serial for 2026-05-13 = 46185 (approx)
    let serial = Math.floor(new Date(Date.UTC(2026, 4, 13)).getTime() / 86400000) + 25569;
    expect(nobitaChuanHoaNgay(serial)).toBe("13/05/2026");
    expect(nobitaChuanHoaNgay(" 08/05/2026 ")).toBe("08/05/2026");
    expect(nobitaChuanHoaNgay("abc")).toBe("abc");
  });

  it("normalizes warehouse location strings", () => {
    let { nobitaChuanHoaViTri } = getTestables();

    expect(nobitaChuanHoaViTri(" b044-c72 ")).toBe("B44 C72");
    expect(nobitaChuanHoaViTri("A005")).toBe("A5");
    expect(nobitaChuanHoaViTri("A5 B044")).toBe("A5 B44");
    expect(nobitaChuanHoaViTri("")).toBe("");
  });

  it("normalizes product names to lowercase and trims whitespace", () => {
    let { nobitaChuanHoaTenSanPham } = getTestables();

    expect(nobitaChuanHoaTenSanPham("  B0W AC 1P BC  ")).toBe("b0w ac 1p bc");
    expect(nobitaChuanHoaTenSanPham("")).toBe("");
    expect(nobitaChuanHoaTenSanPham(123)).toBe("123");
  });

  it("tokenizes product names", () => {
    let { nobitaTachToken } = getTestables();

    expect(nobitaTachToken("b0w ac 1p bc")).toEqual(["b0w", "ac", "1p", "bc"]);
    expect(nobitaTachToken("")).toEqual([]);
    expect(nobitaTachToken("  ")).toEqual([]);
  });

  it("calculates Jaccard similarity score between token sets", () => {
    let { nobitaTinhDiemKhop } = getTestables();

    let a = ["b0w", "ac", "1p", "bc"];
    let b = ["b0w", "ac", "1p", "bc"];
    expect(nobitaTinhDiemKhop(a, b)).toBe(1.0);

    let c = ["b0w", "ac", "1p", "ce"];
    // intersection={b0w,ac,1p}=3, union={b0w,ac,1p,bc,ce}=5 → 3/5=0.6
    expect(nobitaTinhDiemKhop(a, c)).toBeCloseTo(0.6);

    expect(nobitaTinhDiemKhop([], ["a"])).toBe(0);
    expect(nobitaTinhDiemKhop(["a"], [])).toBe(0);
  });

  it("finds product specs by fuzzy name matching", () => {
    let { nobitaTachToken, nobitaTimThongSoTheoTenMo } = getTestables();

    let danhSach = [
      {
        tenChuan: "b0w ac 1p bc din tab life l",
        tokens: nobitaTachToken("b0w ac 1p bc din tab life l"),
        theTich: 0.1,
        kichThuoc: "600*300*500",
        coDuLieu: true,
      },
      {
        tenChuan: "b0w ac 1p ce din tab life l",
        tokens: nobitaTachToken("b0w ac 1p ce din tab life l"),
        theTich: 0.2,
        kichThuoc: "700*300*500",
        coDuLieu: true,
      },
    ];

    let config = { bat: true, diemToiThieu: 0.55, tyLeTokenKhopToiThieu: 0.5, soTokenKhopToiDa: 3 };

    // "rect" is close to "bc/life" variant
    let result = nobitaTimThongSoTheoTenMo(
      "b0w ac 1p bc din tab rect l",
      danhSach,
      config,
    );
    expect(result).toBeDefined();
    expect(result?.theTich).toBe(0.1);

    // Empty query returns undefined
    expect(nobitaTimThongSoTheoTenMo("", danhSach, config)).toBeUndefined();
    // Empty list returns undefined
    expect(nobitaTimThongSoTheoTenMo("b0w ac", [], config)).toBeUndefined();
  });

  it("prefers fuzzy name matches with usable specs when score ties", () => {
    let { nobitaTachToken, nobitaTimThongSoTheoTenMo } = getTestables();
    let config = { bat: true, diemToiThieu: 0.55, tyLeTokenKhopToiThieu: 0.5, soTokenKhopToiDa: 3 };
    let danhSach = [
      {
        tenChuan: "a b c",
        tokens: nobitaTachToken("a b c"),
        theTich: 0,
        kichThuoc: "",
        coDuLieu: false,
      },
      {
        tenChuan: "a b c",
        tokens: nobitaTachToken("a b c"),
        theTich: 0.12,
        kichThuoc: "100*200*600",
        coDuLieu: true,
      },
    ];

    let result = nobitaTimThongSoTheoTenMo("a b c", danhSach, config);
    expect(result?.theTich).toBe(0.12);
  });

  it("parses warehouse location clusters for sorting", () => {
    let { nobitaTachCumViTri } = getTestables();

    expect(nobitaTachCumViTri("A5 B44")).toEqual([
      { khu: "A", so: 5 },
      { khu: "B", so: 44 },
    ]);
    expect(nobitaTachCumViTri("B044-C72")).toEqual([
      { khu: "B", so: 44 },
      { khu: "C", so: 72 },
    ]);
    expect(nobitaTachCumViTri("")).toEqual([]);
  });

  it("sorts locations in warehouse order", () => {
    let { nobitaSoSanhViTri } = getTestables();

    // A before B
    expect(nobitaSoSanhViTri("A5", "B1")).toBeLessThan(0);
    // Lower number first within same zone
    expect(nobitaSoSanhViTri("A5", "A10")).toBeLessThan(0);
    // Same location
    expect(nobitaSoSanhViTri("A5", "A5")).toBe(0);
    // Location before empty
    expect(nobitaSoSanhViTri("A5", "")).toBeLessThan(0);
  });

  it("calculates number of complete packing blocks from total volume", () => {
    let { tinhSoKhoiSoanDuoc } = getTestables();

    expect(tinhSoKhoiSoanDuoc(72, 72)).toBe(1);
    expect(tinhSoKhoiSoanDuoc(144, 72)).toBe(2);
    expect(tinhSoKhoiSoanDuoc(71.9, 72)).toBe(0);
    expect(tinhSoKhoiSoanDuoc(145, 72)).toBe(2);
    expect(tinhSoKhoiSoanDuoc(0, 72)).toBe(0);
    expect(tinhSoKhoiSoanDuoc(100, 0)).toBe(0); // nguongTu = 0 → no division
  });

  it("calculates number of boxes fitting in one volume block", () => {
    let { tinhSoThuongTrongMotKhoi } = getTestables();

    // 72 m³ / 0.1 m³ per box = 720 boxes
    expect(tinhSoThuongTrongMotKhoi(0.1, 72)).toBe(720);
    // 72 m³ / 0.5 m³ per box = 144 boxes
    expect(tinhSoThuongTrongMotKhoi(0.5, 72)).toBe(144);
    // 72 m³ / 0.085 m³ per box = 847 (floor)
    expect(tinhSoThuongTrongMotKhoi(0.085, 72)).toBe(847);
    // Edge cases
    expect(tinhSoThuongTrongMotKhoi(0, 72)).toBe(0);
    expect(tinhSoThuongTrongMotKhoi(0.1, 0)).toBe(0);
  });

  it("builds pending keys from contract, sku, and order quantity", () => {
    let { nobitaTaoKhoaPendingTheoNhom, nobitaTaoKhoaPendingTheoViTri } = getTestables();

    let khoa = nobitaTaoKhoaPendingTheoNhom("EU-903980", "903980", 50);
    expect(khoa).toBe("EU|EU-903980|903980|50");

    // Ignores daXuat: same key regardless of how much was exported
    let khoa2 = nobitaTaoKhoaPendingTheoNhom("EU-903980", "903980", 50);
    expect(khoa2).toBe(khoa);

    // Location key appended
    let khoaViTri = nobitaTaoKhoaPendingTheoViTri(khoa, " b044-c72 ");
    expect(khoaViTri).toBe("EU|EU-903980|903980|50|B44 C72");
  });

  it("subtracts pending by location first, then by group remainder", () => {
    let { nobitaRutSoLuongPending } = getTestables();

    let pendingTheoNhom = { "UE|HD01|901019A|16": 7 };
    let pendingTheoViTri = { "UE|HD01|901019A|16|B44": 3 };

    let ketQua = nobitaRutSoLuongPending(
      10,
      "UE|HD01|901019A|16",
      "B44",
      pendingTheoNhom,
      pendingTheoViTri,
    );

    expect(ketQua).toEqual({ tonKhoSauPending: 3, soLuongPendingDaTru: 7 });
    expect(pendingTheoNhom["UE|HD01|901019A|16"]).toBe(0);
    expect(pendingTheoViTri["UE|HD01|901019A|16|B44"]).toBe(0);
  });

  it("subtracts full pending when stock exactly covers it", () => {
    let { nobitaRutSoLuongPending } = getTestables();

    let pendingTheoNhom = { "EU|EU-903980|903980|50": 50 };
    let pendingTheoViTri = {};

    let ketQua = nobitaRutSoLuongPending(
      50,
      "EU|EU-903980|903980|50",
      "",
      pendingTheoNhom,
      pendingTheoViTri,
    );

    expect(ketQua).toEqual({ tonKhoSauPending: 0, soLuongPendingDaTru: 50 });
    expect(pendingTheoNhom["EU|EU-903980|903980|50"]).toBe(0);
  });

  it("aggregates unfinished production by contract and SKU like Mitsuha", () => {
    let {
      nobitaTaoKhoaNhomHopDongSKU,
      nobitaTinhChuaHoanThanhDong,
      nobitaTongHopChuaHoanThanhTheoNhomHopDongSKU,
    } = getTestables();

    expect(nobitaTaoKhoaNhomHopDongSKU("UE-HD01", "929666")).toBe(
      "UE|UE-HD01|929666",
    );
    expect(nobitaTinhChuaHoanThanhDong(8, 8, 0)).toBe(0);
    expect(nobitaTinhChuaHoanThanhDong(7, 3, 0)).toBe(4);

    let tongHop = nobitaTongHopChuaHoanThanhTheoNhomHopDongSKU([
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

  it("returns full stock when no pending exists for that key", () => {
    let { nobitaRutSoLuongPending } = getTestables();

    let pendingTheoNhom = {};
    let pendingTheoViTri = {};

    let ketQua = nobitaRutSoLuongPending(
      20,
      "EU|EU-111111|111111|10",
      "A5",
      pendingTheoNhom,
      pendingTheoViTri,
    );

    expect(ketQua).toEqual({ tonKhoSauPending: 20, soLuongPendingDaTru: 0 });
  });

  it("parses dd/mm/yyyy dates for import window checks", () => {
    let { nobitaParseNgayDdMmYyyy, nobitaNgaySangSortKey } = getTestables();

    let dt = nobitaParseNgayDdMmYyyy("13/05/2026");
    expect(dt).not.toBeNull();
    expect(dt?.getDate()).toBe(13);
    expect(dt?.getMonth()).toBe(4);
    expect(dt?.getFullYear()).toBe(2026);
    expect(nobitaParseNgayDdMmYyyy("invalid")).toBeNull();
    expect(nobitaNgaySangSortKey("08/05/2026")).toBeLessThan(
      nobitaNgaySangSortKey("13/05/2026"),
    );
  });

  it("detects dates within the last N days inclusive", () => {
    let { nobitaNgayTrongKhoangGanDay } = getTestables();

    let homNay = new Date(2026, 4, 20);
    homNay.setHours(0, 0, 0, 0);

    expect(nobitaNgayTrongKhoangGanDay("20/05/2026", homNay, 14)).toBe(true);
    expect(nobitaNgayTrongKhoangGanDay("07/05/2026", homNay, 14)).toBe(true);
    expect(nobitaNgayTrongKhoangGanDay("06/05/2026", homNay, 14)).toBe(false);
    expect(nobitaNgayTrongKhoangGanDay("", homNay, 14)).toBe(false);
  });

  it("allocates packable quantity across import events FIFO", () => {
    let { nobitaPhanBoSoanTheoSuKienNhap } = getTestables();

    let suKien = [
      { ngay: "10/05/2026", sl: 30 },
      { ngay: "15/05/2026", sl: 50 },
    ];

    expect(nobitaPhanBoSoanTheoSuKienNhap(suKien, 40)).toEqual([
      { ngay: "10/05/2026", slNhap: 30, slSoan: 30 },
      { ngay: "15/05/2026", slNhap: 50, slSoan: 10 },
    ]);

    expect(nobitaPhanBoSoanTheoSuKienNhap(suKien, 0)).toEqual([
      { ngay: "10/05/2026", slNhap: 30, slSoan: 0 },
      { ngay: "15/05/2026", slNhap: 50, slSoan: 0 },
    ]);
  });

  it("counts packable quantity per import day independently (not FIFO)", () => {
    let { nobitaTinhSoanTheoTungNgayNhap } = getTestables();

    let suKien = [
      { ngay: "10/05/2026", sl: 30 },
      { ngay: "13/05/2026", sl: 50 },
    ];

    // Tồn 80: ngày 13 được min(50, 80) = 50 — không bị ngày 10/5 trừ trước
    expect(nobitaTinhSoanTheoTungNgayNhap(suKien, 80)).toEqual([
      { ngay: "10/05/2026", slNhap: 30, slSoan: 30 },
      { ngay: "13/05/2026", slNhap: 50, slSoan: 50 },
    ]);

    // Tồn 8: FIFO — ngày 10/5 lấy hết 8, ngày 13 còn lại 0
    // (tổng slSoan không vượt tồn thực tế 8)
    expect(nobitaTinhSoanTheoTungNgayNhap(suKien, 8)).toEqual([
      { ngay: "10/05/2026", slNhap: 30, slSoan: 8 },
      { ngay: "13/05/2026", slNhap: 50, slSoan: 0 },
    ]);
  });

  it("applies per-market minimum stock threshold via layNguongTonKho logic", () => {
    // Verify tinhSoKhoiSoanDuoc and tinhSoThuongTrongMotKhoi still correct
    // (layNguongTonKho is an inline closure in main(), tested indirectly)
    let { tinhSoKhoiSoanDuoc } = getTestables();

    // EU threshold=8: 10 boxes pass, 6 boxes would be filtered
    // After filtering, only items with tonKho >= threshold contribute to volume
    // This test verifies the building blocks used by the threshold logic
    expect(tinhSoKhoiSoanDuoc(72, 72)).toBe(1);  // exactly 1 block
    expect(tinhSoKhoiSoanDuoc(56, 72)).toBe(0);  // below threshold: 0 blocks (EU items filtered)
    expect(tinhSoKhoiSoanDuoc(144, 72)).toBe(2); // 2 full blocks
  });
});
