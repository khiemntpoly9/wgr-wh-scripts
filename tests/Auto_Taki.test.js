import { beforeAll, describe, expect, it } from "vitest";

import { createRecordingSheet } from "./helpers/excelMocks.js";

const getTestables = () => {
  let testables = globalThis.__AUTO_TAKI_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_Taki test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Auto_Taki.ts");
});

describe("Auto_Taki helpers", () => {
  it("normalizes dates and parses optional numbers", () => {
    let { chuanHoaNgay, ngayToSerial, parseOptionalNumber } = getTestables();

    expect(chuanHoaNgay(ngayToSerial("08/05/2026"))).toBe("08/05/2026");
    expect(chuanHoaNgay(" 09/05/2026 ")).toBe("09/05/2026");
    expect(parseOptionalNumber("")).toBe(0);
    expect(parseOptionalNumber(" 15.5 ")).toBe(15.5);
    expect(parseOptionalNumber("abc")).toBeNull();
  });

  it("detects paired A/B SKUs", () => {
    let { tachSKUCapAB } = getTestables();

    expect(tachSKUCapAB("901019A")).toEqual({
      laSKUCapAB: true,
      skuGoc: "901019",
      hauTo: "A",
    });
    expect(tachSKUCapAB(" 901019b ")).toEqual({
      laSKUCapAB: true,
      skuGoc: "901019",
      hauTo: "B",
    });
    expect(tachSKUCapAB("901019")).toEqual({
      laSKUCapAB: false,
      skuGoc: "",
      hauTo: "",
    });
    expect(tachSKUCapAB("9123456A")).toEqual({
      laSKUCapAB: true,
      skuGoc: "9123456",
      hauTo: "A",
    });
    expect(tachSKUCapAB(9123456)).toEqual({
      laSKUCapAB: false,
      skuGoc: "",
      hauTo: "",
    });
  });

  it("batches contiguous target updates by column and date flag", () => {
    let { ghiTargetUpdatesTheoLo, ngayToSerial } = getTestables();
    let mock = createRecordingSheet();

    ghiTargetUpdatesTheoLo(mock.sheet, [
      { row: 2, col: "A", val: 10 },
      { row: 3, col: "A", val: 20 },
      { row: 5, col: "A", val: 30 },
      { row: 2, col: "B", val: "08/05/2026", isDate: true },
      { row: 3, col: "B", val: "09/05/2026", isDate: true },
      { row: 5, col: "B", val: "text", isDate: false },
    ]);

    expect(mock.calls).toEqual([
      {
        address: "A2:A3",
        values: [[10], [20]],
        numberFormat: null,
        numberFormatLocal: null,
      },
      {
        address: "A5:A5",
        values: [[30]],
        numberFormat: null,
        numberFormatLocal: null,
      },
      {
        address: "B2:B3",
        values: [[ngayToSerial("08/05/2026")], [ngayToSerial("09/05/2026")]],
        numberFormat: null,
        numberFormatLocal: null,
      },
      {
        address: "B5:B5",
        values: [["text"]],
        numberFormat: null,
        numberFormatLocal: null,
      },
      {
        address: "B2:B3",
        values: null,
        numberFormat: "dd/mm/yyyy",
        numberFormatLocal: null,
      },
    ]);
  });
});
