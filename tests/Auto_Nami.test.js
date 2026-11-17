import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_NAMI_TESTABLES__;
  if (!testables) {
    throw new Error("Auto_Nami test helpers were not loaded.");
  }
  return testables;
};

beforeAll(async () => {
  await import("../finished-goods/Auto_Nami.ts");
});

describe("Auto_Nami date parsing helpers", () => {
  it("parses Excel serial date into yyyy-mm-dd key", () => {
    let { namiExcelSerialSangNgayKey } = getTestables();
    let serial =
      Math.floor(new Date(Date.UTC(2026, 4, 9)).getTime() / 86400000) + 25569;

    expect(namiExcelSerialSangNgayKey(serial)).toBe("2026-05-09");
  });

  it("parses text dates in dd/mm/yyyy and yyyy-mm-dd formats", () => {
    let { namiTextSangNgayKey } = getTestables();

    expect(namiTextSangNgayKey("09/05/2026")).toBe("2026-05-09");
    expect(namiTextSangNgayKey("2026-05-09")).toBe("2026-05-09");
    expect(namiTextSangNgayKey("09.05.2026 08:30")).toBe("2026-05-09");
  });

  it("prefers raw cell value when both text and value exist", () => {
    let { namiLayNgayKeyTuO } = getTestables();
    let serial =
      Math.floor(new Date(Date.UTC(2026, 4, 9)).getTime() / 86400000) + 25569;

    // textDate may be locale-formatted/ambiguous; valueDate serial should win.
    expect(namiLayNgayKeyTuO("5/9/2026", serial)).toBe("2026-05-09");
  });
});

describe("Auto_Nami report column order", () => {
  const macDinh = ["thiTruong", "maPhanLoai", "soLuong", "theTich"];

  it("keeps configured order and drops unknown or duplicate columns", () => {
    let { namiChuanHoaThuTuCot } = getTestables();

    expect(
      namiChuanHoaThuTuCot(
        ["soLuong", " thiTruong ", "khongCo", "soLuong"],
        macDinh,
      ),
    ).toEqual(["soLuong", "thiTruong"]);
  });

  it("falls back to default order when config is empty or invalid", () => {
    let { namiChuanHoaThuTuCot } = getTestables();

    expect(namiChuanHoaThuTuCot([], macDinh)).toEqual(macDinh);
    expect(namiChuanHoaThuTuCot(undefined, macDinh)).toEqual(macDinh);
    expect(namiChuanHoaThuTuCot(["sai"], macDinh)).toEqual(macDinh);
  });
});

describe("Auto_Nami location normalization", () => {
  it("normalizes warehouse location clusters", () => {
    let { namiChuanHoaViTri } = getTestables();

    expect(namiChuanHoaViTri("A05  B044")).toBe("A5 B44");
    expect(namiChuanHoaViTri("c12")).toBe("C12");
    expect(namiChuanHoaViTri("")).toBe("");
  });

  it("prefers cell value when displayed text is empty", () => {
    let { namiLayGiaTriO, namiChuanHoaViTri, namiCotChuSangIndex } =
      getTestables();

    expect(namiCotChuSangIndex("I")).toBe(8);
    expect(namiLayGiaTriO(["", ""], ["", "A05  B044"], 1)).toBe("A05  B044");
    expect(
      namiChuanHoaViTri(namiLayGiaTriO(["", ""], ["", "A05  B044"], 1)),
    ).toBe("A5 B44");
  });
});
