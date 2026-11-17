import { beforeAll, describe, expect, it } from "vitest";

const createChainProxy = () => {
  const proxy = new Proxy(function () {}, {
    get: () => proxy,
    apply: () => proxy,
  });
  return proxy;
};

beforeAll(async () => {
  globalThis.ExcelScript = createChainProxy();
  await import("../finished-goods/Auto_Zoro.ts");
});

describe("Auto_Zoro location filter", () => {
  it("normalizes configured location codes and reports invalid ones", () => {
    const { zoroChuanHoaDanhSachViTri } = globalThis.__AUTO_ZORO_TESTABLES__;
    expect(zoroChuanHoaDanhSachViTri([" e023 ", "E23", "E23E24", "#$", ""])).toEqual({
      codes: ["E23", "E24"],
      ignored: ["#$"],
    });
    expect(zoroChuanHoaDanhSachViTri([])).toEqual({ codes: [], ignored: [] });
  });

  it("matches a location cell against any configured code segment", () => {
    const { zoroKhopViTri } = globalThis.__AUTO_ZORO_TESTABLES__;
    expect(zoroKhopViTri("E23 E24", ["E24"])).toBe(true);
    expect(zoroKhopViTri("e023", ["E23"])).toBe(true);
    expect(zoroKhopViTri("E23, F10", ["F10"])).toBe(true);
    expect(zoroKhopViTri("E230", ["E23"])).toBe(false);
    expect(zoroKhopViTri("", ["E23"])).toBe(false);
    expect(zoroKhopViTri("E23", [])).toBe(false);
  });

  it("routes rows at the requested locations to the location sheet whatever their contract", () => {
    const { zoroXacDinhDich } = globalThis.__AUTO_ZORO_TESTABLES__;
    const codes = ["E23", "E24"];
    expect(zoroXacDinhDich("UEM5003", "E23", ["UE"], codes)).toEqual({ thiTruong: "UE", viTri: true });
    expect(zoroXacDinhDich("JPM5003", "E23 F10", ["UE"], codes)).toEqual({ thiTruong: "", viTri: true });
    expect(zoroXacDinhDich("", "e024", ["UE"], codes)).toEqual({ thiTruong: "", viTri: true });
    expect(zoroXacDinhDich("uwm5003", "F10", ["UW"], codes)).toEqual({ thiTruong: "UW", viTri: false });
    expect(zoroXacDinhDich("AUM5003", "", ["UE"], codes)).toEqual({ thiTruong: "", viTri: false });
    expect(zoroXacDinhDich("UEM5003", "E23", ["UE"], [])).toEqual({ thiTruong: "UE", viTri: false });
  });
});
