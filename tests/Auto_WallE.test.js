import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_WALLE_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_WallE test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Auto_WallE.ts");
});

describe("Auto_WallE helpers", () => {
  it("normalizes SKU, location, name and market code helpers", () => {
    let {
      lamSach,
      chuanHoaSKU6So,
      chuanHoaViTri,
      chuanHoaTen,
      chuanHoaMaThiTruong2KyTu,
      laMaThiTruongRutGon,
    } = getTestables();

    expect(lamSach("  HD \u00A0\u200B 01  ")).toBe("HD01");
    expect(chuanHoaSKU6So("SKU-123456-A")).toBe("123456");
    expect(chuanHoaSKU6So("12345")).toBe("");
    expect(chuanHoaViTri("a01 / b002-c0003")).toBe("A1 B2 C3");
    expect(chuanHoaTen("  Tên \u00A0 Hàng  A  ")).toBe("tên hàng a");
    expect(chuanHoaMaThiTruong2KyTu(" ue-01 ")).toBe("UE");
    expect(laMaThiTruongRutGon("EU")).toBe(true);
    expect(laMaThiTruongRutGon("EU1")).toBe(false);
  });

  it("prefers the closest sufficient candidate before fallback ordering", () => {
    let { chonUngVienPhuHop } = getTestables();

    let danhSach = [
      {
        hopDong: "HD03",
        maHang: "901001",
        tenHang: "A",
        kichThuoc: "",
        slDonHang: 20,
        daXuat: 0,
        deXuatXuat: 9,
        trongLuong: 1,
        theTich: 1,
        viTri: "A1",
        tonKho: 11,
        dangSoanCong: 2,
        conLaiDon: 20,
      },
      {
        hopDong: "HD01",
        maHang: "901001",
        tenHang: "B",
        kichThuoc: "",
        slDonHang: 20,
        daXuat: 0,
        deXuatXuat: 6,
        trongLuong: 1,
        theTich: 1,
        viTri: "A2",
        tonKho: 8,
        dangSoanCong: 1,
        conLaiDon: 20,
      },
      {
        hopDong: "HD02",
        maHang: "901001",
        tenHang: "C",
        kichThuoc: "",
        slDonHang: 20,
        daXuat: 0,
        deXuatXuat: 6,
        trongLuong: 1,
        theTich: 1,
        viTri: "A3",
        tonKho: 7,
        dangSoanCong: 0,
        conLaiDon: 20,
      },
    ];

    expect(chonUngVienPhuHop(danhSach, 5)?.hopDong).toBe("HD02");
    expect(chonUngVienPhuHop(danhSach, 12)?.hopDong).toBe("HD03");
    expect(chonUngVienPhuHop(danhSach, 0)?.hopDong).toBe("HD03");
  });

  it("matches product specs by fuzzy token overlap on product names", () => {
    let { chuanHoaTen, tachTokenTenSanPham, tinhDiemKhopTenMo, timThongSoTheoTenMo } =
      getTestables();

    let tenTraCuu = chuanHoaTen("B0W ac 1p bc din tab rect I");
    let tokensTraCuu = tachTokenTenSanPham(tenTraCuu);

    let tenBc = chuanHoaTen("B0W ac 1p bc din tab life l");
    let tenCe = chuanHoaTen("B0W ac 1p ce din tab life l");
    let tenWt = chuanHoaTen("B0W ac 1p wt din tab life l");

    expect(tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenBc))).toBeGreaterThan(
      tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenCe)),
    );
    expect(tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenBc))).toBeGreaterThan(
      tinhDiemKhopTenMo(tokensTraCuu, tachTokenTenSanPham(tenWt)),
    );

    let danhSach = [tenBc, tenCe, tenWt].map((tenChuan, i) => ({
      tenChuan,
      tokens: tachTokenTenSanPham(tenChuan),
      thongSo: { kichThuoc: `size-${i}`, trongLuong: 1 + i, theTich: 0.1 },
      coDuLieu: true,
    }));

    let ketQua = timThongSoTheoTenMo(tenTraCuu, danhSach);
    expect(ketQua?.kichThuoc).toBe("size-0");
  });
});
