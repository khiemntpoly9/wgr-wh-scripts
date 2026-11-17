import { beforeAll, describe, expect, it } from "vitest";

const createChainProxy = () => {
  const proxy = new Proxy(function () {}, {
    get: () => proxy,
    apply: () => proxy,
  });
  return proxy;
};

const createPlan = (name, values) => ({
  name,
  clear_existing: true,
  writes: [{ address: `A1:M${values.length}`, values }],
  formulas: [],
  styles: [],
  border_ranges: [],
  merges: [],
  auto_filter: null,
  freeze_rows: 1,
  summary: { market_code: name.slice(0, 2), total_lines: values.length - 1, total_stock: 0 },
});

const createWorkbook = (failingSheet) => {
  const writes = [];
  const createSheet = (name) => {
    const chain = createChainProxy();
    return new Proxy(
      {},
      {
        get(_, prop) {
          if (prop !== "getRange") return () => chain;
          return (address) =>
            new Proxy(
              {},
              {
                get(__, method) {
                  if (method !== "setValues") return () => chain;
                  return (values) => {
                    if (name === failingSheet) throw new Error("Timeout ghi sheet");
                    writes.push({ sheet: name, address, rows: values.length });
                  };
                },
              },
            );
        },
      },
    );
  };
  const sheets = {};
  return {
    writes,
    workbook: {
      getWorksheet: (name) => sheets[name],
      addWorksheet: (name) => (sheets[name] = createSheet(name)),
    },
  };
};

beforeAll(async () => {
  globalThis.ExcelScript = createChainProxy();
  await import("../finished-goods/Auto_Zoro_Super.ts");
});

describe("Auto_Zoro_Super API helpers", () => {
  it("builds the backend payload without an empty report date", () => {
    const { zoroSuperTaoPayload } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const payload = zoroSuperTaoPayload({
      apiUrl: "https://example.com",
      apiToken: "token",
      marketCodes: [" eu ", "AU"],
      reportDate: "",
      sheetNamePrefix: "Tồn kho",
      receiptDateFilter: { days: [14], months: [9], years: [2026] },
    });
    expect(payload.market_codes).toEqual(["EU", "AU"]);
    expect(payload.report_date).toBeUndefined();
    expect(payload.receipt_date_filter.years).toEqual([2026]);
  });

  it("sends trimmed location codes only when configured", () => {
    const { zoroSuperTaoPayload } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const base = {
      apiUrl: "https://example.com",
      apiToken: "token",
      marketCodes: ["UE"],
      reportDate: "",
      sheetNamePrefix: "Tồn kho",
      receiptDateFilter: { days: [], months: [], years: [] },
      debugSku: "",
    };
    expect(
      zoroSuperTaoPayload({ ...base, locationCodes: [" E23 ", "", "E24"] }).location_codes,
    ).toEqual(["E23", "E24"]);
    expect(zoroSuperTaoPayload({ ...base, locationCodes: [] })).not.toHaveProperty("location_codes");
    expect(zoroSuperTaoPayload({ ...base, locationCodes: [" "] })).not.toHaveProperty("location_codes");
  });

  it("summarizes a location sheet by its name", () => {
    const { zoroSuperGhiCacSheet } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const { workbook } = createWorkbook(null);
    const plan = createPlan("Vị trí-Tồn kho-2026-11-13", [["STT"], [1], [2]]);
    plan.summary = { location_codes: ["E23"], total_lines: 2, total_stock: 5 };
    const result = zoroSuperGhiCacSheet(workbook, [plan], "");
    expect(result.tomTat).toEqual(["Vị trí-Tồn kho-2026-11-13: 2 dòng, tồn 5."]);
  });

  it("describes location filter results from the response summary", () => {
    const { zoroSuperMoTaViTri } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    expect(zoroSuperMoTaViTri({ location_codes: ["E23", "E24"], ignored_location_codes: ["#$"] })).toEqual([
      "Vị trí lọc: E23, E24",
      "Bỏ qua mã vị trí không hợp lệ: #$",
    ]);
    expect(zoroSuperMoTaViTri({ location_codes: [], ignored_location_codes: ["#$"] })).toEqual([
      "Vị trí lọc: (không có mã hợp lệ)",
      "Bỏ qua mã vị trí không hợp lệ: #$",
    ]);
    expect(zoroSuperMoTaViTri({ total_sheets: 1 })).toEqual([]);
    expect(zoroSuperMoTaViTri(undefined)).toEqual([]);
  });

  it("validates URL, ISO date and response format", () => {
    const { zoroSuperKiemTraUrl, zoroSuperLaNgayIso, zoroSuperKiemTraKeHoach } =
      globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    expect(zoroSuperKiemTraUrl("https://app.wgr-wh.site/api/v1/test")).toBeNull();
    expect(zoroSuperKiemTraUrl("http://localhost/api/test")).toContain("HTTPS");
    expect(zoroSuperLaNgayIso("2026-09-14")).toBe(true);
    expect(zoroSuperLaNgayIso("2026-02-29")).toBe(false);
    expect(zoroSuperKiemTraKeHoach({ format: "auto_zoro_backend_v1", sheets: [] })).toBe(true);
    expect(zoroSuperKiemTraKeHoach({ format: "wrong", sheets: [] })).toBe(false);
  });

  it("batches contiguous single-cell formulas into setFormulas ranges", () => {
    const { zoroSuperParseAddress, zoroSuperGhiFormulas } =
      globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    expect(zoroSuperParseAddress("I12")).toEqual({ col: "I", row: 12 });
    expect(zoroSuperParseAddress("A1:M10")).toBeNull();

    const calls = [];
    const sheet = {
      getRange(address) {
        return {
          setFormulas(matrix) {
            calls.push({ address, matrix });
          },
          setFormula(formula) {
            calls.push({ address, formula });
          },
        };
      },
    };
    zoroSuperGhiFormulas(sheet, [
      { address: "A2", formula: "=ROW()-1" },
      { address: "A3", formula: "=ROW()-1" },
      { address: "A5", formula: "=ROW()-1" },
      { address: "I2", formula: "=H2*G2" },
    ]);
    expect(calls).toEqual([
      {
        address: "A2:A3",
        matrix: [["=ROW()-1"], ["=ROW()-1"]],
      },
      { address: "A5:A5", matrix: [["=ROW()-1"]] },
      { address: "I2:I2", matrix: [["=H2*G2"]] },
    ]);
  });

  it("writes large value blocks in row chunks", () => {
    const { zoroSuperGhiGiaTri } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const calls = [];
    const sheet = {
      getRange(address) {
        return { setValues: (values) => calls.push({ address, rows: values.length, first: values[0][0] }) };
      },
    };
    const values = Array.from({ length: 2501 }, (_, index) => [index + 1, "x"]);
    zoroSuperGhiGiaTri(sheet, "A1:B2501", values, 1000);
    expect(calls).toEqual([
      { address: "A1:B1000", rows: 1000, first: 1 },
      { address: "A1001:B2000", rows: 1000, first: 1001 },
      { address: "A2001:B2501", rows: 501, first: 2001 },
    ]);

    calls.length = 0;
    zoroSuperGhiGiaTri(sheet, "Sheet!A1", [["a"]], 1000);
    expect(calls).toEqual([{ address: "Sheet!A1", rows: 1, first: "a" }]);
  });

  it("finds debug SKU rows in a sheet plan, matching text or numeric cells", () => {
    const { zoroSuperTimDongSku } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const plan = createPlan("UE-Tồn kho-2026-11-13", [
      ["STT", "Mã đơn hàng", "Mã hợp đồng", "SKU"],
      [1, "WGM25003A", "UEM5003", "936771", "DL7", 145, 145],
      [2, "WGM25003A", "UEM5004", 912345],
      [3, "WGM25003A", "UEM5005", 936771, "DL7", 3, 3],
    ]);
    expect(zoroSuperTimDongSku(plan, " 936771 ")).toEqual([
      "dòng 2: 1 | WGM25003A | UEM5003 | 936771 | DL7 | 145 | 145",
      "dòng 4: 3 | WGM25003A | UEM5005 | 936771 | DL7 | 3 | 3",
    ]);
    expect(zoroSuperTimDongSku(plan, "")).toEqual([]);
  });

  it("keeps writing the remaining sheets when one sheet fails and reports the failure", () => {
    const { zoroSuperGhiCacSheet } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const { workbook, writes } = createWorkbook("UE-Tồn kho-2026-11-13");
    const result = zoroSuperGhiCacSheet(
      workbook,
      [
        createPlan("UE-Tồn kho-2026-11-13", [["STT"], [1]]),
        createPlan("UW-Tồn kho-2026-11-13", [["STT"], [1], [2]]),
      ],
      "",
    );
    expect(writes).toEqual([{ sheet: "UW-Tồn kho-2026-11-13", address: "A1:M3", rows: 3 }]);
    expect(result.loi).toEqual(['"UE-Tồn kho-2026-11-13": Error: Timeout ghi sheet']);
    expect(result.tomTat).toEqual(["UW: 2 dòng, tồn 0."]);
  });

  it("replaces an enabled filter and freezes the header row", () => {
    const { zoroSuperGhiCacSheet } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const calls = [];
    const sheet = {
      getUsedRange: () => undefined,
      getRange: () => createChainProxy(),
      getAutoFilter: () => ({
        getEnabled: () => true,
        remove: () => calls.push("filter.remove"),
        apply: (address) => calls.push(`filter.apply ${address}`),
      }),
      getFreezePanes: () => ({
        unfreeze: () => calls.push("freeze.unfreeze"),
        freezeRows: (rows) => calls.push(`freeze.freezeRows ${rows}`),
      }),
    };
    const workbook = { getWorksheet: () => sheet, addWorksheet: () => sheet };
    const plan = createPlan("UE-Tồn kho-2026-11-13", [["STT"], [1], [2]]);
    plan.auto_filter = "A1:M3";
    zoroSuperGhiCacSheet(workbook, [plan], "");
    expect(calls).toEqual([
      "filter.remove",
      "filter.apply A1:M3",
      "freeze.unfreeze",
      "freeze.freezeRows 1",
    ]);
  });

  it("logs debug SKU rows per sheet, including sheets without a match", () => {
    const { zoroSuperGhiCacSheet } = globalThis.__AUTO_ZORO_SUPER_TESTABLES__;
    const { workbook } = createWorkbook(null);
    const result = zoroSuperGhiCacSheet(
      workbook,
      [
        createPlan("UE-Tồn kho-2026-11-13", [["STT", "HĐ", "SKU"], [1, "UEM5003", "936771"]]),
        createPlan("UW-Tồn kho-2026-11-13", [["STT", "HĐ", "SKU"]]),
      ],
      "936771",
    );
    expect(result.debug).toEqual([
      'DEBUG SKU 936771 @ "UE-Tồn kho-2026-11-13": dòng 2: 1 | UEM5003 | 936771',
      'DEBUG SKU 936771 @ "UW-Tồn kho-2026-11-13": không có dòng.',
    ]);
  });
});
