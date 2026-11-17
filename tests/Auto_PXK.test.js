import { beforeAll, describe, expect, it } from "vitest";

import {
  createMappedSheet,
  createMockExcelScriptEnums,
  createMockRange,
  createWorkbookWithSheets,
} from "./helpers/excelMocks.js";

const getTestables = () => {
  let testables = globalThis.__AUTO_PXK_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_PXK test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  globalThis.ExcelScript = createMockExcelScriptEnums();

  await import("../Auto_PXK.ts");
});

describe("Auto_PXK helpers", () => {
  it("maps string alignment values to ExcelScript enums", () => {
    let { toHorizontalAlignment, toVerticalAlignment } = getTestables();

    expect(toHorizontalAlignment("center")).toBe("center");
    expect(toHorizontalAlignment("right")).toBe("right");
    expect(toVerticalAlignment("middle")).toBe("center");
    expect(toVerticalAlignment("bottom")).toBe("bottom");
  });

  it("finds the header row by scanning column A", () => {
    let { findHeaderRow } = getTestables();
    let sheet = createMappedSheet({
      usedRange: createMockRange({ rowCount: 4 }),
      ranges: {
        "A1:A4": createMockRange({
          values: [["noise"], ["STT"], ["sku only"], ["sku stt name"]],
        }),
      },
    });

    expect(findHeaderRow(sheet, ["STT", "SKU"])).toBe(3);
    expect(findHeaderRow(sheet, ["NOT_FOUND"])).toBe(-1);
  });

  it("reads footer signatures and finds footer start rows", () => {
    let {
      getFooterSignatureFromWorkbook,
      findFooterStartRowFromSignature,
    } = getTestables();

    let workbook = createWorkbookWithSheets({
      footer: {
        getUsedRange() {
          return createMockRange({
            intersection: createMockRange({
              values: [["", "ABC", "Footer signature long enough", ""]],
            }),
          });
        },
      },
    });

    let sheet = createMappedSheet({
      ranges: {
        "A2:Q6": createMockRange({
          values: [
            ["row", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            [
              "footer signature long enough here",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
            ],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
          ],
        }),
      },
    });

    expect(getFooterSignatureFromWorkbook(workbook)).toBe(
      "Footer signature long enough",
    );
    expect(
      findFooterStartRowFromSignature(
        sheet,
        2,
        6,
        "Footer signature long enough",
      ),
    ).toBe(3);
  });

  it("limits data bounds before footer and ignores blank rows inside the range", () => {
    let { findDataBounds } = getTestables();

    let workbook = createWorkbookWithSheets({
      footer: {
        getUsedRange() {
          return createMockRange({
            intersection: createMockRange({
              values: [["", "Footer signature long enough", ""]],
            }),
          });
        },
      },
    });

    let emptyRow = () => Array(17).fill("");
    let sheet = createMappedSheet({
      usedRange: createMockRange({ rowIndex: 0, rowCount: 10 }),
      ranges: {
        "A2:Q10": createMockRange({
          values: [
            emptyRow(),
            emptyRow(),
            emptyRow(),
            emptyRow(),
            emptyRow(),
            [
              "footer signature long enough starts here",
              ...Array(16).fill(""),
            ],
            emptyRow(),
            emptyRow(),
            emptyRow(),
          ],
        }),
        "D2:D6": createMockRange({
          values: [["item 1"], [""], ["item 2"], [""], ["item 3"]],
        }),
      },
    });

    expect(findDataBounds(sheet, 2, workbook, [], "Q")).toEqual({
      startRow: 2,
      lastRow: 6,
    });
  });

  it("finds footer when signature is not in column A (e.g. after STT overwrote A)", () => {
    let { findFooterStartRowFromSignature } = getTestables();

    let sheet = createMappedSheet({
      ranges: {
        "A2:C4": createMockRange({
          values: [
            ["=ROW()-1", "", ""],
            ["", "", "Footer signature long enough"],
            ["", "", ""],
          ],
        }),
      },
    });

    expect(
      findFooterStartRowFromSignature(
        sheet,
        2,
        4,
        "Footer signature long enough",
        "C",
      ),
    ).toBe(3);
  });

  it("limits data bounds before existing summary row on re-run", () => {
    let { findDataBounds } = getTestables();

    let sheet = createMappedSheet({
      usedRange: createMockRange({ rowIndex: 0, rowCount: 8 }),
      ranges: {
        "D2:D4": createMockRange({
          values: [["item 1"], ["item 2"], [""]],
        }),
        "D5:D8": createMockRange({
          values: [
            [""],
            ["=SUM would be row 6 — not data"],
            ["looks like product but is below summary"],
            [""],
          ],
        }),
        "F2:F8": createMockRange({
          values: [
            [1],
            [2],
            [""],
            [""],
            [0],
            [""],
            [""],
          ],
          formulas: [
            [""],
            [""],
            [""],
            [""],
            ["=SUM(F2:F3)"],
            [""],
            [""],
          ],
        }),
      },
    });

    expect(findDataBounds(sheet, 2, createWorkbookWithSheets({}), ["F"])).toEqual(
      {
        startRow: 2,
        lastRow: 3,
      },
    );
  });

  it("parses third dimension from size strings", () => {
    let { getDimensionThirdValue, isPackagePackUnit } = getTestables();

    expect(getDimensionThirdValue("44*24*1,1")).toBe(1.1);
    expect(getDimensionThirdValue("25*45*1.1")).toBe(1.1);
    expect(getDimensionThirdValue("50*21,5*21,5")).toBe(21.5);
    expect(getDimensionThirdValue("44*24*25")).toBe(25);
    expect(getDimensionThirdValue("44*24")).toBe(null);

    expect(isPackagePackUnit("44*24*1,1")).toBe(true);
    expect(isPackagePackUnit("25x45x1,1")).toBe(true);
    expect(isPackagePackUnit("50*21,5*21,5")).toBe(false);
    expect(isPackagePackUnit("44*24*25")).toBe(false);
  });

  it("resolves target rows when copying data from pre-packing-list to sale", () => {
    let { resolvePxkDataCopyPlan } = getTestables();

    expect(resolvePxkDataCopyPlan(10, { startRow: 11, lastRow: 15 })).toEqual({
      targetStartRow: 11,
      targetLastRow: 15,
    });
    expect(resolvePxkDataCopyPlan(8, { startRow: 20, lastRow: 22 })).toEqual({
      targetStartRow: 9,
      targetLastRow: 11,
    });
    expect(resolvePxkDataCopyPlan(5, null)).toBe(null);
    expect(resolvePxkDataCopyPlan(5, { startRow: 12, lastRow: 11 })).toBe(null);
  });
});

describe("Auto_PXK container volume", () => {
  const volumeConfig = (targetTotalVolume) => ({
    enabled: true,
    targetTotalVolume,
    productCol: "C",
    dimCol: "E",
    qtyCol: "I",
    unitVolumeCol: "N",
    totalVolumeCol: "O",
    divisor: 1000000,
    adjustSide: "all",
    dimensionDecimals: 0,
    maxChangePercent: 10,
    wholeNumberTarget: true,
    tolerance: 0.01,
  });

  const createVolumeSheet = (rows) => {
    let cols = { C: [], E: [], I: [], N: [] };
    for (let row of rows) {
      cols.C.push([row.product]);
      cols.E.push([row.dim]);
      cols.I.push([row.qty]);
      cols.N.push([row.unitVolume]);
    }
    let writes = [];
    let sheet = {
      getRange(address) {
        let col = address.match(/[A-Z]+/)[0];
        return {
          getValues: () => cols[col].map((row) => [...row]),
          getFormulas: () => cols[col].map(() => [""]),
          setValues(values) {
            cols[col] = values;
            writes.push({ address, type: "values", values });
          },
          setFormulas(values) {
            writes.push({ address, type: "formulas", values });
          },
        };
      },
    };
    return { sheet, writes };
  };

  const lineVolume = (dim, qty) => {
    let parts = String(dim)
      .split("*")
      .map((part) => Number(part.replace(",", ".")));
    let unit = Math.round(((parts[0] * parts[1] * parts[2]) / 1000000) * 1000) / 1000;
    return Math.round(unit * qty * 100) / 100;
  };

  const runAdjust = (rows, target) => {
    let { autoAdjustPxkContainerVolume } = getTestables();
    let { sheet, writes } = createVolumeSheet(rows);
    let logs = [];
    let originalLog = console.log;
    console.log = (...args) => {
      logs.push(args.join(" "));
    };
    try {
      autoAdjustPxkContainerVolume(
        sheet,
        volumeConfig(target),
        2,
        1 + rows.length,
        2 + rows.length,
      );
    } finally {
      console.log = originalLog;
    }
    let dimWrite = writes.find((write) => write.type === "values");
    let formulaWrite = writes.find((write) => write.type === "formulas");
    return {
      dims: dimWrite.values.map((row) => row[0]),
      formulas: formulaWrite.values.map((row) => row[0]),
      logs,
    };
  };

  it("tăng kích thước theo từng cm nguyên, không còn bước thập phân", () => {
    let { dims, formulas } = runAdjust(
      [
        { product: "A1", dim: "50*40.5*30", qty: 300, unitVolume: 0.061 },
        { product: "A1b", dim: "50*40,5*30", qty: 0, unitVolume: 0.061 },
      ],
      19,
    );

    expect(dims[0]).toMatch(/^\d+\*\d+\*\d+$/);
    expect(dims[0]).toBe(dims[1]);
    expect(formulas[0]).toBe(`=ROUND((${dims[0]})/1000000, 3)`);
    expect(formulas[0].includes(".")).toBe(false);
    let total = lineVolume(dims[0], 300) + lineVolume(dims[1], 0);
    expect(total).toBeGreaterThanOrEqual(19);
    expect(total).toBeLessThanOrEqual(19.99);
  });

  it("giảm kích thước theo từng cm nguyên để vào khoảng số khối nguyên", () => {
    let { dims } = runAdjust(
      [{ product: "B1", dim: "50*40*30", qty: 300, unitVolume: 0.06 }],
      17,
    );

    expect(dims[0]).toMatch(/^\d+\*\d+\*\d+$/);
    let [length, width, height] = dims[0].split("*").map(Number);
    expect(Number.isInteger(length - 50)).toBe(true);
    expect(Number.isInteger(width - 40)).toBe(true);
    expect(Number.isInteger(height - 30)).toBe(true);
    expect(dims[0]).not.toBe("50*40*30");
    let total = lineVolume(dims[0], 300);
    expect(total).toBeGreaterThanOrEqual(17);
    expect(total).toBeLessThanOrEqual(17.99);
  });

  it("không thay đổi kích thước hàng PACKAGE (chiều cao 1.1)", () => {
    let { dims, formulas } = runAdjust(
      [
        { product: "B1", dim: "50*40*30", qty: 300, unitVolume: 0.06 },
        { product: "P1", dim: "2*6*1.1", qty: 100, unitVolume: 0 },
        { product: "P2", dim: "44*24*1,1", qty: 50, unitVolume: 0.001 },
      ],
      17,
    );

    expect(dims[0]).not.toBe("50*40*30");
    expect(dims[1]).toBe("2*6*1.1");
    expect(dims[2]).toBe("44*24*1,1");
    expect(formulas[1]).toBe("");
    expect(formulas[2]).toBe("");
    let total = lineVolume(dims[0], 300) + 0 + Math.round(0.001 * 50 * 100) / 100;
    expect(total).toBeGreaterThanOrEqual(17);
    expect(total).toBeLessThanOrEqual(17.99);
  });
});
