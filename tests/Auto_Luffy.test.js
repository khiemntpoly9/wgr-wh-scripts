import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_LUFFY_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_Luffy test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Auto_Luffy.ts");
});

describe("Auto_Luffy pure helpers", () => {
  it("luffyLamSach removes spaces and zero-width characters", () => {
    let { luffyLamSach } = getTestables();

    expect(luffyLamSach("  AB \u00A0\u200B 12  ")).toBe("AB12");
    expect(luffyLamSach(1230)).toBe("1230");
  });

  it("luffyHangNguonTrong treats empty contract, sku and quantity as blank row", () => {
    let { luffyHangNguonTrong } = getTestables();

    expect(luffyHangNguonTrong("", "", "")).toBe(true);
    expect(luffyHangNguonTrong("  ", "\u00A0", " \u200B ")).toBe(true);
    expect(luffyHangNguonTrong("", "", 0)).toBe(false);
    expect(luffyHangNguonTrong("HD01", "", "")).toBe(false);
    expect(luffyHangNguonTrong("", "SKU01", "")).toBe(false);
  });

  it("luffyLaySoStt only accepts positive integer sequence values", () => {
    let { luffyLaySoStt } = getTestables();

    expect(luffyLaySoStt(1)).toBe(1);
    expect(luffyLaySoStt(" 002 ")).toBe(2);
    expect(luffyLaySoStt(3.1)).toBeNull();
    expect(luffyLaySoStt("1A")).toBeNull();
    expect(luffyLaySoStt(0)).toBeNull();
    expect(luffyLaySoStt("")).toBeNull();
  });

  it("luffyTrongSoLuong keeps numeric zero as a valid quantity", () => {
    let { luffyTrongSoLuong } = getTestables();

    expect(luffyTrongSoLuong(0)).toBe(false);
    expect(luffyTrongSoLuong(12)).toBe(false);
    expect(luffyTrongSoLuong("")).toBe(true);
    expect(luffyTrongSoLuong("   ")).toBe(true);
  });

  it("luffyXuLyCotCSkuSoBo10So0 trims at most ten trailing zeroes", () => {
    let { luffyXuLyCotCSkuSoBo10So0 } = getTestables();

    expect(luffyXuLyCotCSkuSoBo10So0(false, 12000)).toBe(12000);
    expect(luffyXuLyCotCSkuSoBo10So0(true, 12000)).toBe(12);
    expect(luffyXuLyCotCSkuSoBo10So0(true, "450000")).toBe(45);
    expect(luffyXuLyCotCSkuSoBo10So0(true, 12345)).toBe(12345);
    expect(luffyXuLyCotCSkuSoBo10So0(true, "ABC")).toBe("ABC");
    expect(luffyXuLyCotCSkuSoBo10So0(true, true)).toBe(true);
    expect(luffyXuLyCotCSkuSoBo10So0(true, 1000000000000)).toBe(100);
  });

  it("luffyTachKichThuoc splits dài*rộng*cao with comma or dot decimals", () => {
    let { luffyTachKichThuoc } = getTestables();

    expect(luffyTachKichThuoc("113,9*58,7*13,1")).toEqual([113.9, 58.7, 13.1]);
    expect(luffyTachKichThuoc("23*23*50")).toEqual([23, 23, 50]);
    expect(luffyTachKichThuoc(" 10,5 * 2 * 3 ")).toEqual([10.5, 2, 3]);
    expect(luffyTachKichThuoc("")).toEqual(["", "", ""]);
    expect(luffyTachKichThuoc("only-one")).toEqual(["only-one", "", ""]);
  });
});
