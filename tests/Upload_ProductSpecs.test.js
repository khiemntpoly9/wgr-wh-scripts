import { beforeAll, describe, expect, it } from "vitest";

const helpers = () => {
  const value = globalThis.__UPLOAD_PRODUCT_SPECS_TESTABLES__;
  if (!value) throw new Error("Upload_ProductSpecs helpers were not loaded.");
  return value;
};

beforeAll(async () => {
  await import("../Upload_ProductSpecs.ts");
});

describe("Upload_ProductSpecs", () => {
  it("parses Vietnamese dates, ISO dates and Excel serial dates", () => {
    const { productSpecsParseDateKey, productSpecsFormatDateKey } = helpers();
    const expected = productSpecsParseDateKey("06/09/2026");

    expect(productSpecsFormatDateKey(expected)).toBe("06/09/2026");
    expect(productSpecsParseDateKey("2026-09-06")).toBe(expected);
    expect(productSpecsParseDateKey(46271)).toBe(expected);
    expect(productSpecsParseDateKey("31/02/2026")).toBeNull();
  });

  it("calculates Monday-Sunday weeks and full months", () => {
    const { productSpecsParseDateKey, productSpecsFormatDateKey, productSpecsGetPeriod } = helpers();
    const anchor = productSpecsParseDateKey("06/09/2026");

    const week = productSpecsGetPeriod("week", anchor);
    expect(productSpecsFormatDateKey(week.startKey)).toBe("31/08/2026");
    expect(productSpecsFormatDateKey(week.endKey)).toBe("06/09/2026");

    const month = productSpecsGetPeriod("month", anchor);
    expect(productSpecsFormatDateKey(month.startKey)).toBe("01/09/2026");
    expect(productSpecsFormatDateKey(month.endKey)).toBe("30/09/2026");
  });

  it("filters the period, computes m3 and keeps the latest row for each SKU", () => {
    const { productSpecsParseDateKey, productSpecsReadRows } = helpers();
    const rows = [
      ["Ngày", "SKU", "Tên", "", "Kg", "Dài", "Rộng", "Cao"],
      ["31/08/2026", "SKU-1", "Cũ", "", 10, 100, 50, 40],
      ["01/09/2026", "sku-1", "Mới 1", "", "12,5", 100, 50, 40],
      ["03/09/2026", "SKU-1", "Mới 2", "", 13, 120, 50, 40],
      ["04/09/2026", "SKU-2", "Hàng 2", "", 8, 80, 40, 30],
    ];

    const result = productSpecsReadRows(
      rows,
      productSpecsParseDateKey("01/09/2026"),
      productSpecsParseDateKey("07/09/2026"),
    );

    expect(result.errors).toEqual([]);
    expect(result.items).toHaveLength(2);
    expect(result.duplicateCount).toBe(1);
    expect(result.items[0]).toMatchObject({
      sku: "SKU-1",
      name: "Mới 2",
      unit_weight: 13,
      unit_volume: 0.24,
    });
  });

  it("reports invalid in-period rows without blocking valid SKUs", () => {
    const { productSpecsParseDateKey, productSpecsReadRows } = helpers();
    const rows = [
      ["Ngày", "SKU", "Tên", "", "Kg", "Dài", "Rộng", "Cao"],
      ["02/09/2026", "SKU-1", "", "", 10, 100, 50, 40],
      ["02/09/2026", "SKU-2", "Hàng 2", "", 8, 80, 40, 30],
    ];

    const result = productSpecsReadRows(
      rows,
      productSpecsParseDateKey("02/09/2026"),
      productSpecsParseDateKey("02/09/2026"),
    );

    expect(result.items).toHaveLength(1);
    expect(result.errors[0]).toContain("Dòng 2");
    expect(result.errors[0]).toContain("tên thành phẩm (C)");
  });

  it("uses 1.1 cm when the height column is blank", () => {
    const { productSpecsParseDateKey, productSpecsReadRows } = helpers();
    const rows = [
      ["Ngày", "SKU", "Tên", "", "Kg", "Dài", "Rộng", "Cao"],
      ["02/09/2026", "SKU-1", "Hàng 1", "", 10, 100, 50, ""],
    ];

    const dateKey = productSpecsParseDateKey("02/09/2026");
    const result = productSpecsReadRows(rows, dateKey, dateKey);

    expect(result.errors).toEqual([]);
    expect(result.items[0]).toMatchObject({
      height: 1.1,
      unit_volume: 0.0055,
    });
  });
});
