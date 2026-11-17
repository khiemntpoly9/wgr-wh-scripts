import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_DORAEMON_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_Doraemon test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Auto_Doraemon.ts");
});

describe("Auto_Doraemon helpers", () => {
  it("cleans values and parses Excel-style numbers", () => {
    let { lamSach, parseSoTuExcel } = getTestables();

    expect(lamSach("  HD \u00A0\u200B 01 ")).toBe("HD01");
    expect(parseSoTuExcel("1.234,56")).toBe(1234.56);
    expect(parseSoTuExcel("1,234.56")).toBe(1234.56);
    expect(parseSoTuExcel("abc")).toBe(0);
  });

  it("builds pending keys using contract, sku, order quantity, exported quantity, and normalized location", () => {
    let { chuanHoaViTri, taoKhoaPendingTheoNhom, taoKhoaPendingTheoViTri } =
      getTestables();

    expect(chuanHoaViTri(" b044-c72 ")).toBe("B44 C72");

    let khoaNhom = taoKhoaPendingTheoNhom(" ue-01 ", " 929666 ", 16, 8);
    expect(khoaNhom).toBe("UE|ue-01|929666|16|8");
    expect(taoKhoaPendingTheoViTri(khoaNhom, " b044-c72 ")).toBe(
      "UE|ue-01|929666|16|8|B44 C72",
    );
  });

  it("subtracts pending by location first, then by group remainder", () => {
    let { rutSoLuongPending } = getTestables();

    let pendingTheoNhom = {
      "UE|HD01|929666|16|8": 4,
    };
    let pendingTheoViTri = {
      "UE|HD01|929666|16|8|B44": 3,
    };

    let ketQua = rutSoLuongPending(
      10,
      "UE|HD01|929666|16|8",
      "B44",
      pendingTheoNhom,
      pendingTheoViTri,
    );

    expect(ketQua).toEqual({
      tonKhoSauPending: 3,
      soLuongPendingDaTru: 7,
    });
    expect(pendingTheoNhom["UE|HD01|929666|16|8"]).toBe(0);
    expect(pendingTheoViTri["UE|HD01|929666|16|8|B44"]).toBe(0);
  });
});
