import { beforeAll, describe, expect, it, vi } from "vitest";

let T;
beforeAll(async () => {
  await import("../fabric/Auto_Reco_Stock_Check.ts");
  T = globalThis.__AUTO_RECO_STOCK_CHECK_TESTABLES__;
});

const col = (values) => values.map((v) => [v]);

describe("Auto_Reco_Stock_Check config and normalization", () => {
  it("exposes expected RECO_STOCK_CHECK_CONFIG and 8-column headers", () => {
    expect(T.RECO_STOCK_CHECK_CONFIG).toMatchObject({
      sourceSheet: "RECO_DATA",
      targetSheet: "RECO_STOCK_CHECK",
      startRow: 8,
      columns: {
        date: "C",
        type: "D",
        itemCode: "F",
        qtyMeters: "I",
      },
      headerRow: 5,
      startDataRow: 6,
      templateInputRowCount: 10,
    });
    expect(T.RECO_STOCK_CHECK_HEADERS).toEqual([
      "STT",
      "Mã vật tư (Item Code)",
      "Tổng Nhập IN (m)",
      "Tổng Xuất OUT (m)",
      "Số lượng tồn còn lại (m)",
      "Trạng thái tồn",
      "Ngày nhập cuối",
      "Ngày xuất cuối",
    ]);
  });

  it("normalizes numeric, scientific-notation, and string item codes cleanly", () => {
    expect(T.recoStockNormalizeItemCode(110702000516)).toBe("110702000516");
    expect(T.recoStockNormalizeItemCode("1.10702000516E+11")).toBe("110702000516");
    expect(T.recoStockNormalizeItemCode("1,10702000516E+11")).toBe("110702000516");
    expect(T.recoStockNormalizeItemCode("201050001.0")).toBe("201050001");
    expect(T.recoStockNormalizeItemCode(" 110702000516\u00A0")).toBe("110702000516");
    expect(T.recoStockNormalizeItemCode("\u200B  ")).toBe("");
    expect(T.recoStockNormalizeItemCode(null)).toBe("");
    expect(T.recoStockNormalizeItemCode(true)).toBe("");
  });

  it("converts safe integer item codes to JS number while keeping non-numeric or leading-zero codes as string", () => {
    expect(T.recoStockToCellItemCode("110702000516")).toBe(110702000516);
    expect(T.recoStockToCellItemCode("FAB-001")).toBe("FAB-001");
    expect(T.recoStockToCellItemCode("00123")).toBe("00123");
    expect(T.recoStockToCellItemCode("")).toBe("");
  });

  it("extracts user item codes from column B, deduplicates, and stops at TỔNG CỘNG", () => {
    expect(
      T.recoStockExtractUserItemCodes(
        col([
          110702000516,
          "",
          " 201050001 ",
          "1.10702000516E+11",
          "FAB-001",
          "",
          "TỔNG CỘNG",
          999999999,
        ]),
      ),
    ).toEqual(["110702000516", "201050001", "FAB-001"]);

    expect(
      T.recoStockExtractUserItemCodes(col(["", "", "TỔNG CỘNG"])),
    ).toEqual([]);
  });
});

describe("Auto_Reco_Stock_Check parsing, aggregation, and status resolution", () => {
  const makeSampleColumns = () => ({
    date: col([
      "01/11/2026",
      "05/11/2026",
      "02/11/2026",
      "03/11/2026",
      "04/11/2026",
      "06/11/2026",
      "07/11/2026",
      "08/11/2026",
    ]),
    type: col(["IN", "OUT", "IN", "OUT", "IN", "OUT", "ADJUST", "IN"]),
    itemCode: col([
      110702000516,
      "1.10702000516E+11",
      "201050001.0",
      201050001,
      "MAT-ZERO",
      "MAT-ZERO",
      110702000516, // invalid txn type
      999999999, // non-target item
    ]),
    qtyMeters: col(["1,250.5", "250,25", 500, "620,5", 100, 100, 50, 20]),
  });

  it("resolves stock status accurately with floating-point rounding", () => {
    expect(T.recoStockResolveStatus(100.25)).toBe("Còn hàng");
    expect(T.recoStockResolveStatus(0)).toBe("Hết hàng");
    expect(T.recoStockResolveStatus(0.00000001)).toBe("Hết hàng");
    expect(T.recoStockResolveStatus(-0.01)).toBe("Âm kho");
  });

  it("aggregates IN, OUT, balance, status, and separate lastInDate / lastOutDate per item", () => {
    const targets = ["110702000516", "201050001", "MAT-ZERO", "MAT-UNUSED"];
    const txns = T.recoStockParseRows(makeSampleColumns(), targets);
    expect(txns).toHaveLength(6);

    const summaries = T.recoStockAggregateByItem(txns, targets);
    expect(summaries).toEqual([
      {
        itemCode: "110702000516",
        totalIn: 1250.5,
        totalOut: 250.25,
        balance: 1000.25,
        status: "Còn hàng",
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        lastInDate: "2026-11-01",
        lastOutDate: "2026-11-05",
      },
      {
        itemCode: "201050001",
        totalIn: 500,
        totalOut: 620.5,
        balance: -120.5,
        status: "Âm kho",
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        lastInDate: "2026-11-02",
        lastOutDate: "2026-11-03",
      },
      {
        itemCode: "MAT-ZERO",
        totalIn: 100,
        totalOut: 100,
        balance: 0,
        status: "Hết hàng",
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        lastInDate: "2026-11-04",
        lastOutDate: "2026-11-06",
      },
      {
        itemCode: "MAT-UNUSED",
        totalIn: 0,
        totalOut: 0,
        balance: 0,
        status: "Hết hàng",
        inCount: 0,
        outCount: 0,
        totalCount: 0,
        lastInDate: null,
        lastOutDate: null,
      },
    ]);

    expect(T.recoStockSummarizeStatusCounts(summaries)).toEqual({
      inStock: 1,
      outOfStock: 2,
      negative: 1,
    });
  });

  it("tracks the latest IN date and latest OUT date across multiple transactions", () => {
    const txns = T.recoStockParseRows(
      {
        date: col(["10/11/2026", "01/11/2026", "15/11/2026", "08/11/2026"]),
        type: col(["IN", "IN", "OUT", "OUT"]),
        itemCode: col(["FAB-001", "FAB-001", "FAB-001", "FAB-001"]),
        qtyMeters: col([50, 100, 30, 20]),
      },
      ["FAB-001"],
    );

    const [summary] = T.recoStockAggregateByItem(txns, ["FAB-001"]);
    expect(summary).toMatchObject({
      totalIn: 150,
      totalOut: 50,
      balance: 100,
      status: "Còn hàng",
      lastInDate: "2026-11-10",
      lastOutDate: "2026-11-15",
    });
  });

  it("parses contiguous C:I sheet matrix via recoStockParseSheetRows", () => {
    const matrix = [
      ["01/11/2026", "IN", "ignoredE", 110702000516, "ignoredG", "ignoredH", "1,250.5"],
      ["02/11/2026", " out ", "", "1.10702000516E+11", "", "", "250,25"],
      ["03/11/2026", "IN", "", 999999999, "", "", 100],
      ["04/11/2026", "ADJUST", "", 110702000516, "", "", 50],
    ];
    const txns = T.recoStockParseSheetRows(matrix, ["110702000516"], 8);
    expect(txns).toEqual([
      {
        rowNumber: 8,
        dateIso: "2026-11-01",
        type: "IN",
        itemCode: "110702000516",
        qtyMeters: 1250.5,
      },
      {
        rowNumber: 9,
        dateIso: "2026-11-02",
        type: "OUT",
        itemCode: "110702000516",
        qtyMeters: 250.25,
      },
    ]);
  });

  it("builds stock check rows with Excel formulas, status formulas, date serials, and TỔNG CỘNG row", () => {
    const targets = ["110702000516", "201050001", "MAT-ZERO"];
    const txns = T.recoStockParseRows(makeSampleColumns(), targets);
    const summaries = T.recoStockAggregateByItem(txns, targets);
    const rows = T.recoStockBuildRows(summaries, 6);

    expect(rows).toHaveLength(4);
    expect(rows[0]).toEqual([
      1,
      110702000516,
      1250.5,
      250.25,
      "=C6-D6",
      '=IF(ROUND(E6,6)>0,"Còn hàng",IF(ROUND(E6,6)<0,"Âm kho","Hết hàng"))',
      T.recoStockIsoToSerial("2026-11-01"),
      T.recoStockIsoToSerial("2026-11-05"),
    ]);
    expect(rows[1]).toEqual([
      2,
      201050001,
      500,
      620.5,
      "=C7-D7",
      '=IF(ROUND(E7,6)>0,"Còn hàng",IF(ROUND(E7,6)<0,"Âm kho","Hết hàng"))',
      T.recoStockIsoToSerial("2026-11-02"),
      T.recoStockIsoToSerial("2026-11-03"),
    ]);
    expect(rows[2]).toEqual([
      3,
      "MAT-ZERO",
      100,
      100,
      "=C8-D8",
      '=IF(ROUND(E8,6)>0,"Còn hàng",IF(ROUND(E8,6)<0,"Âm kho","Hết hàng"))',
      T.recoStockIsoToSerial("2026-11-04"),
      T.recoStockIsoToSerial("2026-11-06"),
    ]);
    expect(rows[3]).toEqual([
      "",
      "TỔNG CỘNG",
      "=SUM(C6:C8)",
      "=SUM(D6:D8)",
      "=SUM(E6:E8)",
      "Còn: 1 | Hết: 1 | Âm: 1",
      "",
      "",
    ]);
  });

  it("builds empty template rows with TỔNG CỘNG formulas", () => {
    const rows = T.recoStockBuildEmptyTemplateRows(10, 6);
    expect(rows).toHaveLength(11);
    expect(rows[0]).toEqual([1, "", "", "", "", "", "", ""]);
    expect(rows[9]).toEqual([10, "", "", "", "", "", "", ""]);
    expect(rows[10]).toEqual([
      "",
      "TỔNG CỘNG",
      "=SUM(C6:C15)",
      "=SUM(D6:D15)",
      "=SUM(E6:E15)",
      "",
      "",
      "",
    ]);
  });

  it("detects last data row via getUsedRange(true) and diagnoses source rows", () => {
    const emptySheet = { getUsedRange: vi.fn(() => undefined) };
    expect(T.recoStockGetLastDataRow(emptySheet)).toBe(0);

    const dataSheet = {
      getUsedRange: vi.fn(() => ({
        getLastRow: () => ({ getRowIndex: () => 24 }),
      })),
    };
    expect(T.recoStockGetLastDataRow(dataSheet)).toBe(25);

    const diag = T.recoStockDiagnoseSourceRows(
      [
        ["01/11/2026", "IN", "", 110702000516, "", "", 100],
        ["02/11/2026", "ADJUST", "", 110702000516, "", "", 50],
        ["03/11/2026", "OUT", "", 999999999, "", "", 20],
        ["04/11/2026", "IN", "", "", "", "", 10],
      ],
      ["110702000516"],
    );
    expect(diag).toEqual({
      totalRows: 4,
      nonEmptyItemRows: 3,
      matchedItemRows: 2,
      validTxnRows: 1,
    });
  });
});
