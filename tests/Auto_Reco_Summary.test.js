import { beforeAll, describe, expect, it, vi } from "vitest";

let T;
beforeAll(async () => {
  await import("../fabric/Auto_Reco_Summary.ts");
  T = globalThis.__AUTO_RECO_SUMMARY_TESTABLES__;
});

const col = (values) => values.map((v) => [v]);

describe("Auto_Reco_Summary config and normalization", () => {
  it("exposes expected RECO_DATA config defaults without hardcoded targetItemCodes", () => {
    expect(T.RECO_SUMMARY_CONFIG).toMatchObject({
      sourceSheet: "RECO_DATA",
      targetSheet: "RECO_SUMMARY",
      startRow: 8,
      columns: {
        date: "C",
        type: "D",
        itemCode: "F",
        qtyMeters: "I",
      },
      summaryHeaderRow: 5,
      summaryStartRow: 6,
      templateInputRowCount: 10,
    });
    expect(T.RECO_SUMMARY_CONFIG.targetItemCodes).toBeUndefined();
  });

  it("normalizes numeric, scientific-notation, and string item codes cleanly", () => {
    expect(T.recoNormalizeItemCode(110702000516)).toBe("110702000516");
    expect(T.recoNormalizeItemCode(201050001)).toBe("201050001");
    expect(T.recoNormalizeItemCode(1e21)).toBe("1000000000000000000000");
    expect(T.recoNormalizeItemCode("1.10702000516E+11")).toBe("110702000516");
    expect(T.recoNormalizeItemCode("1,10702000516E+11")).toBe("110702000516");
    expect(T.recoNormalizeItemCode("1.10702000516e+11")).toBe("110702000516");
    expect(T.recoNormalizeItemCode("2,01050001E+8")).toBe("201050001");
    expect(T.recoNormalizeItemCode("1.107E+5")).toBe("110700");
    expect(T.recoNormalizeItemCode(" 110702000516\u00A0")).toBe("110702000516");
    expect(T.recoNormalizeItemCode("201050001.0")).toBe("201050001");
    expect(T.recoNormalizeItemCode("\u200B  ")).toBe("");
    expect(T.recoNormalizeItemCode(null)).toBe("");
    expect(T.recoNormalizeItemCode(true)).toBe("");
  });

  it("normalizes transaction type IN / OUT", () => {
    expect(T.recoNormalizeTxnType("IN")).toBe("IN");
    expect(T.recoNormalizeTxnType(" in\u00A0")).toBe("IN");
    expect(T.recoNormalizeTxnType("Out ")).toBe("OUT");
    expect(T.recoNormalizeTxnType("RETURN")).toBeNull();
    expect(T.recoNormalizeTxnType("")).toBeNull();
    expect(T.recoNormalizeTxnType(null)).toBeNull();
  });

  it("parses meter quantities with dot/comma decimals and thousands separators", () => {
    expect(T.recoParseNumber(125.5)).toBe(125.5);
    expect(T.recoParseNumber("125,5")).toBe(125.5);
    expect(T.recoParseNumber(" 125.75 ")).toBe(125.75);
    expect(T.recoParseNumber("1,234.56")).toBe(1234.56);
    expect(T.recoParseNumber("1.234,56")).toBe(1234.56);
    expect(T.recoParseNumber("1,234,567")).toBe(1234567);
    expect(T.recoParseNumber("1.234.567")).toBe(1234567);
    expect(T.recoParseNumber("-42,25")).toBe(-42.25);
    expect(T.recoParseNumber("")).toBe(0);
    expect(T.recoParseNumber("abc")).toBe(0);
    expect(T.recoParseNumber(false)).toBe(0);
  });

  it("parses Excel serial dates, Date objects, and formatted date strings", () => {
    const serial = T.recoIsoToSerial("2026-11-17");
    expect(T.recoParseDate(serial)).toBe("2026-11-17");
    expect(T.recoParseDate(String(serial))).toBe("2026-11-17");
    expect(T.recoParseDate(new Date(2026, 10, 17))).toBe("2026-11-17");
    expect(T.recoParseDate("17/11/2026")).toBe("2026-11-17");
    expect(T.recoParseDate("17-11-2026")).toBe("2026-11-17");
    expect(T.recoParseDate("17.11.2026")).toBe("2026-11-17");
    expect(T.recoParseDate("2026-11-17")).toBe("2026-11-17");
    expect(T.recoFormatDateVn("2026-11-17")).toBe("17/11/2026");

    expect(T.recoParseDate("31/02/2026")).toBeNull();
    expect(T.recoParseDate("")).toBeNull();
    expect(T.recoParseDate(0)).toBeNull();
  });
});

describe("Auto_Reco_Summary row parsing and aggregation", () => {
  const makeSampleColumns = () => ({
    date: col([
      "01/11/2026",
      T.recoIsoToSerial("2026-11-02"),
      "03-11-2026",
      "04/11/2026",
      "05/11/2026",
      "06/11/2026",
      "07/11/2026",
    ]),
    type: col(["IN", " out ", "in", "OUT", "IN", "ADJUST", "IN"]),
    itemCode: col([
      110702000516,
      " 110702000516 ",
      "201050001.0",
      201050001,
      999999999, // non-target item
      110702000516, // invalid txn type
      "", // empty item code
    ]),
    qtyMeters: col(["1,250.5", "250,25", 500, "620,5", 100, 50, 20]),
  });

  it("filters rows to target item codes and tracks 1-based row numbers from row 8", () => {
    const txns = T.recoParseRows(makeSampleColumns(), ["110702000516", "201050001"]);
    expect(txns).toHaveLength(4);
    expect(txns[0]).toEqual({
      rowNumber: 8,
      dateIso: "2026-11-01",
      type: "IN",
      itemCode: "110702000516",
      qtyMeters: 1250.5,
    });
    expect(txns[1]).toEqual({
      rowNumber: 9,
      dateIso: "2026-11-02",
      type: "OUT",
      itemCode: "110702000516",
      qtyMeters: 250.25,
    });
    expect(txns[2]).toEqual({
      rowNumber: 10,
      dateIso: "2026-11-03",
      type: "IN",
      itemCode: "201050001",
      qtyMeters: 500,
    });
    expect(txns[3]).toEqual({
      rowNumber: 11,
      dateIso: "2026-11-04",
      type: "OUT",
      itemCode: "201050001",
      qtyMeters: 620.5,
    });
  });

  it("aggregates IN, OUT, Balance, counts, and date range per item code", () => {
    const txns = T.recoParseRows(makeSampleColumns(), ["110702000516", "201050001"]);
    const summaries = T.recoAggregateByItem(txns, ["110702000516", "201050001", "300000001"]);

    expect(summaries).toEqual([
      {
        itemCode: "110702000516",
        totalIn: 1250.5,
        totalOut: 250.25,
        balance: 1000.25,
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        firstDate: "2026-11-01",
        lastDate: "2026-11-02",
      },
      {
        itemCode: "201050001",
        totalIn: 500,
        totalOut: 620.5,
        balance: -120.5,
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        firstDate: "2026-11-03",
        lastDate: "2026-11-04",
      },
      {
        itemCode: "300000001",
        totalIn: 0,
        totalOut: 0,
        balance: 0,
        inCount: 0,
        outCount: 0,
        totalCount: 0,
        firstDate: null,
        lastDate: null,
      },
    ]);
  });

  it("aggregates daily breakdown sorted by item code, date ascending (nulls last), then IN before OUT", () => {
    const txns = T.recoParseRows({
      date: col(["02/11/2026", "01/11/2026", "02/11/2026", "02/11/2026", "", ""]),
      type: col(["OUT", "IN", "IN", "IN", "OUT", "IN"]),
      itemCode: col([201050001, 110702000516, 201050001, 201050001, 110702000516, 110702000516]),
      qtyMeters: col([120, 100, 199.95, 100.05, 15.25, 50]),
    });

    const daily = T.recoAggregateByDateAndItem(txns);
    expect(daily).toEqual([
      {
        dateIso: "2026-11-01",
        itemCode: "110702000516",
        type: "IN",
        qtyMeters: 100,
        runningBalance: 100,
        txnCount: 1,
      },
      {
        dateIso: null,
        itemCode: "110702000516",
        type: "IN",
        qtyMeters: 50,
        runningBalance: 150,
        txnCount: 1,
      },
      {
        dateIso: null,
        itemCode: "110702000516",
        type: "OUT",
        qtyMeters: 15.25,
        runningBalance: 134.75,
        txnCount: 1,
      },
      {
        dateIso: "2026-11-02",
        itemCode: "201050001",
        type: "IN",
        qtyMeters: 300,
        runningBalance: 300,
        txnCount: 2,
      },
      {
        dateIso: "2026-11-02",
        itemCode: "201050001",
        type: "OUT",
        qtyMeters: 120,
        runningBalance: 180,
        txnCount: 1,
      },
    ]);
  });

  it("builds summary and daily sheet rows with numeric item codes, Excel formulas, and date serials", () => {
    const txns = T.recoParseRows(makeSampleColumns(), ["110702000516", "201050001"]);
    const summaries = T.recoAggregateByItem(txns, ["110702000516", "201050001"]);
    const summaryRows = T.recoBuildSummaryRows(summaries, 6);

    expect(summaryRows).toHaveLength(3);
    expect(summaryRows[0]).toEqual([
      1,
      110702000516,
      1250.5,
      250.25,
      "=C6-D6",
      1,
      1,
      "=F6+G6",
      T.recoIsoToSerial("2026-11-01"),
      T.recoIsoToSerial("2026-11-02"),
    ]);
    expect(typeof summaryRows[0][1]).toBe("number");
    expect(summaryRows[1]).toEqual([
      2,
      201050001,
      500,
      620.5,
      "=C7-D7",
      1,
      1,
      "=F7+G7",
      T.recoIsoToSerial("2026-11-03"),
      T.recoIsoToSerial("2026-11-04"),
    ]);
    expect(typeof summaryRows[1][1]).toBe("number");
    expect(summaryRows[2]).toEqual([
      "",
      "TỔNG CỘNG",
      "=SUM(C6:C7)",
      "=SUM(D6:D7)",
      "=SUM(E6:E7)",
      "=SUM(F6:F7)",
      "=SUM(G6:G7)",
      "=SUM(H6:H7)",
      "",
      "",
    ]);

    expect(T.RECO_DAILY_HEADERS).toEqual([
      "STT",
      "Ngày",
      "Mã hàng (Item Code)",
      "Loại (IN/OUT)",
      "Số mét (m)",
      "Số lượng tồn (m)",
      "Số giao dịch",
    ]);

    const daily = T.recoAggregateByDateAndItem(txns);
    const dailyRows = T.recoBuildDailyRows(daily, 11);
    expect(dailyRows).toHaveLength(4);
    expect(dailyRows[0]).toEqual([
      1,
      T.recoIsoToSerial("2026-11-01"),
      110702000516,
      "IN",
      1250.5,
      1250.5,
      1,
    ]);
    expect(dailyRows[1]).toEqual([
      2,
      T.recoIsoToSerial("2026-11-02"),
      110702000516,
      "OUT",
      250.25,
      1000.25,
      1,
    ]);
    expect(dailyRows[2]).toEqual([
      3,
      T.recoIsoToSerial("2026-11-03"),
      201050001,
      "IN",
      500,
      500,
      1,
    ]);
    expect(dailyRows[3]).toEqual([
      4,
      T.recoIsoToSerial("2026-11-04"),
      201050001,
      "OUT",
      620.5,
      -120.5,
      1,
    ]);
    expect(typeof dailyRows[0][2]).toBe("number");
  });

  it("builds separate IN and OUT daily rows grouped by itemCode with independent runningBalance", () => {
    const txns = T.recoParseRows({
      date: col(["05/11/2026", "05/11/2026", "06/11/2026", "07/11/2026", ""]),
      type: col(["OUT", "IN", "IN", "OUT", "OUT"]),
      itemCode: col(["FAB-001", "FAB-001", 110702000516, "FAB-001", 110702000516]),
      qtyMeters: col([45.5, 120.75, 100, 100, 30]),
    });

    const daily = T.recoAggregateByDateAndItem(txns);
    const dailyRows = T.recoBuildDailyRows(daily);
    expect(dailyRows).toEqual([
      [1, T.recoIsoToSerial("2026-11-06"), 110702000516, "IN", 100, 100, 1],
      [2, "", 110702000516, "OUT", 30, 70, 1],
      [3, T.recoIsoToSerial("2026-11-05"), "FAB-001", "IN", 120.75, 120.75, 1],
      [4, T.recoIsoToSerial("2026-11-05"), "FAB-001", "OUT", 45.5, 75.25, 1],
      [5, T.recoIsoToSerial("2026-11-07"), "FAB-001", "OUT", 100, -24.75, 1],
    ]);
  });

  it("parses contiguous C:I sheet matrix via recoParseSheetRows and filters by target item code", () => {
    const matrix = [
      ["01/11/2026", "IN", "ignoredE", 110702000516, "ignoredG", "ignoredH", "1,250.5"],
      ["invalid-date", "OUT", "", 999999999, "", "", "bad-number"],
      ["02/11/2026", " out ", "", "1.10702000516E+11", "", "", "250,25"],
      ["03/11/2026", "IN", "", "", "", "", 100],
      ["04/11/2026", "ADJUST", "", 110702000516, "", "", 50],
    ];
    const txns = T.recoParseSheetRows(matrix, ["110702000516"], 8);
    expect(txns).toEqual([
      {
        rowNumber: 8,
        dateIso: "2026-11-01",
        type: "IN",
        itemCode: "110702000516",
        qtyMeters: 1250.5,
      },
      {
        rowNumber: 10,
        dateIso: "2026-11-02",
        type: "OUT",
        itemCode: "110702000516",
        qtyMeters: 250.25,
      },
    ]);
  });

  it("parses rows with numeric and scientific-notation item codes into target codes", () => {
    const txns = T.recoParseRows(
      {
        date: col(["01/11/2026", "02/11/2026", "03/11/2026"]),
        type: col(["IN", "OUT", "IN"]),
        itemCode: col([110702000516, "1,10702000516E+11", "1.10702000516E+11"]),
        qtyMeters: col([100, 25, 50]),
      },
      ["110702000516"],
    );
    expect(txns).toHaveLength(3);
    expect(txns.every((t) => t.itemCode === "110702000516")).toBe(true);
    const summaries = T.recoAggregateByItem(txns, ["110702000516"]);
    expect(summaries[0]).toMatchObject({
      itemCode: "110702000516",
      totalIn: 150,
      totalOut: 25,
      balance: 125,
      totalCount: 3,
    });
  });
});

describe("Auto_Reco_Summary user-driven input, empty template, and integer formatting", () => {
  it("converts safe integer item codes to JS number while keeping non-numeric or leading-zero codes as string", () => {
    expect(T.recoToCellItemCode("110702000516")).toBe(110702000516);
    expect(T.recoToCellItemCode("201050001")).toBe(201050001);
    expect(T.recoToCellItemCode("1.10702000516E+11")).toBe(110702000516);
    expect(T.recoToCellItemCode("FAB-001")).toBe("FAB-001");
    expect(T.recoToCellItemCode("00123")).toBe("00123");
    expect(T.recoToCellItemCode("")).toBe("");
  });

  it("extracts user item codes from column B, deduplicates, and stops at TỔNG CỘNG or daily section", () => {
    expect(
      T.recoExtractUserItemCodes(
        col([
          110702000516,
          "",
          " 201050001 ",
          "1.10702000516E+11",
          "FAB-001",
          "",
          "TỔNG CỘNG",
          "",
          "Ngày",
          999999999,
        ]),
      ),
    ).toEqual(["110702000516", "201050001", "FAB-001"]);

    expect(
      T.recoExtractUserItemCodes(col(["", "", "", "TỔNG CỘNG"])),
    ).toEqual([]);

    expect(
      T.recoExtractUserItemCodes(
        col([300000001, "2. CHI TIẾT THEO NGÀY & MÃ HÀNG (DAILY BREAKDOWN)", 110702000516]),
      ),
    ).toEqual(["300000001"]);
  });

  it("builds empty template rows with TỔNG CỘNG formulas and no pre-filled item codes", () => {
    const rows = T.recoBuildEmptyTemplateRows(10, 6);
    expect(rows).toHaveLength(11);
    expect(rows[0]).toEqual([1, "", "", "", "", "", "", "", "", ""]);
    expect(rows[9]).toEqual([10, "", "", "", "", "", "", "", "", ""]);
    expect(rows[10]).toEqual([
      "",
      "TỔNG CỘNG",
      "=SUM(C6:C15)",
      "=SUM(D6:D15)",
      "=SUM(E6:E15)",
      "=SUM(F6:F15)",
      "=SUM(G6:G15)",
      "=SUM(H6:H15)",
      "",
      "",
    ]);
  });

  it("detects last data row via getUsedRange(true) and returns 0 on empty sheet", () => {
    const emptySheet = { getUsedRange: vi.fn(() => undefined) };
    expect(T.recoGetLastDataRow(emptySheet)).toBe(0);
    expect(emptySheet.getUsedRange).toHaveBeenCalledWith(true);

    const dataSheet = {
      getUsedRange: vi.fn(() => ({
        getLastRow: () => ({ getRowIndex: () => 24 }),
      })),
    };
    expect(T.recoGetLastDataRow(dataSheet)).toBe(25);
    expect(dataSheet.getUsedRange).toHaveBeenCalledWith(true);
  });

  it("diagnoses source rows for empty item codes, unmatched targets, and invalid txn types", () => {
    const rows = [
      ["01/11/2026", "IN", "", 110702000516, "", "", 100],
      ["02/11/2026", "ADJUST", "", 110702000516, "", "", 50],
      ["03/11/2026", "OUT", "", 999999999, "", "", 20],
      ["04/11/2026", "IN", "", "", "", "", 10],
    ];
    expect(T.recoDiagnoseSourceRows(rows, ["110702000516"])).toEqual({
      totalRows: 4,
      nonEmptyItemRows: 3,
      matchedItemRows: 2,
      validTxnRows: 1,
    });
  });
});
