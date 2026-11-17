import { beforeAll, describe, expect, it } from "vitest";

beforeAll(async () => {
  await import("../Auto_Sanji_Super.ts");
});

describe("Auto_Sanji_Super API helpers", () => {
  it("maps the Office Script configuration to backend field names", () => {
    const { sanjiSuperTaoPayload } = globalThis.__AUTO_SANJI_SUPER_TESTABLES__;
    expect(
      sanjiSuperTaoPayload({
        apiUrl: "https://example.com",
        apiToken: "token",
        marketCodes: [" eu "],
        sheetName: "Tồn kho",
        classificationOrder: ["F", "A"],
        classificationOrderByMarket: { JP: ["T", "M"] },
      }),
    ).toEqual({
      market_codes: ["EU"],
      sheet_name: "Tồn kho",
      classification_order: ["F", "A"],
      classification_order_by_market: { JP: ["T", "M"] },
    });
  });

  it("accepts a workbook plan or an intentional empty result", () => {
    const { sanjiSuperKiemTraUrl, sanjiSuperKiemTraKeHoach } =
      globalThis.__AUTO_SANJI_SUPER_TESTABLES__;
    expect(sanjiSuperKiemTraUrl("https://app.wgr-wh.site/api/v1/test")).toBeNull();
    expect(sanjiSuperKiemTraKeHoach({ format: "auto_sanji_backend_v1", sheet: null })).toBe(true);
    expect(
      sanjiSuperKiemTraKeHoach({
        format: "auto_sanji_backend_v1",
        sheet: { writes: [], formulas: [] },
      }),
    ).toBe(true);
    expect(sanjiSuperKiemTraKeHoach({ format: "wrong", sheet: null })).toBe(false);
  });
});
