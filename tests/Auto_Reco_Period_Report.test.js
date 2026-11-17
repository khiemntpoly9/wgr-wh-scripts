import { beforeAll, describe, expect, it, vi } from "vitest";

let T;
beforeAll(async () => {
  await import("../fabric/Auto_Reco_Period_Report.ts");
  T = globalThis.__AUTO_RECO_PERIOD_REPORT_TESTABLES__;
});

const col = (values) => values.map((v) => [v]);

describe("Auto_Reco_Period_Report config and parsing", () => {
  it("exposes expected RECO_PERIOD_REPORT_CONFIG and 11-column headers", () => {
    expect(T.RECO_PERIOD_REPORT_CONFIG).toMatchObject({
      sourceSheet: "RECO_DATA",
      targetSheet: "RECO_REPORT",
      startRow: 8,
      columns: {
        date: "C",
        type: "D",
        itemCode: "F",
        qtyMeters: "I",
      },
      includeDaily: true,
      includeWeekly: true,
      includeMonthly: true,
    });
    expect(T.RECO_PERIOD_HEADERS).toEqual([
      "STT",
      "Kỳ báo cáo",
      "Từ ngày",
      "Đến ngày",
      "Tổng Nhập IN (m)",
      "Tổng Xuất OUT (m)",
      "Chênh lệch Net (m)",
      "Số lần IN",
      "Số lần OUT",
      "Tổng giao dịch",
      "Số mã hàng",
    ]);
  });

  it("normalizes item codes, txn types, numbers, and dates", () => {
    expect(T.recoReportNormalizeItemCode(110702000516)).toBe("110702000516");
    expect(T.recoReportNormalizeItemCode("1.10702000516E+11")).toBe("110702000516");
    expect(T.recoReportNormalizeItemCode("201050001.0")).toBe("201050001");
    expect(T.recoReportNormalizeItemCode("")).toBe("");

    expect(T.recoReportNormalizeTxnType(" in ")).toBe("IN");
    expect(T.recoReportNormalizeTxnType("Out")).toBe("OUT");
    expect(T.recoReportNormalizeTxnType("ADJUST")).toBeNull();

    expect(T.recoReportParseNumber("1,250.5")).toBe(1250.5);
    expect(T.recoReportParseNumber("250,25")).toBe(250.25);

    const serial = T.recoReportIsoToSerial("2026-11-17");
    expect(T.recoReportParseDate(serial)).toBe("2026-11-17");
    expect(T.recoParseDate).toBeUndefined();
    expect(T.recoReportParseDate("17/11/2026")).toBe("2026-11-17");
    expect(T.recoReportParseDate("2026-11-17")).toBe("2026-11-17");
    expect(T.recoReportParseDate("31/02/2026")).toBeNull();
    expect(T.recoReportFormatDateVn("2026-11-17")).toBe("17/11/2026");
  });

  it("parses C:I sheet matrix and filters by target item codes and date/item options", () => {
    const matrix = [
      ["01/11/2026", "IN", "", 110702000516, "", "", "1,250.5"],
      ["02/11/2026", "OUT", "", 201050001, "", "", "250,25"],
      ["", "IN", "", 110702000516, "", "", 100],
      ["05/11/2026", "ADJUST", "", 110702000516, "", "", 50],
      ["06/11/2026", "IN", "", "", "", "", 50],
    ];
    const txns = T.recoReportParseSheetRows(matrix, [], 8);
    expect(txns).toHaveLength(3);
    expect(txns[0]).toEqual({
      rowNumber: 8,
      dateIso: "2026-11-01",
      type: "IN",
      itemCode: "110702000516",
      qtyMeters: 1250.5,
    });
    expect(txns[2].dateIso).toBeNull();

    const filteredByDate = T.recoReportFilterTransactions(txns, {
      fromDate: "2026-11-01",
      toDate: "2026-11-02",
    });
    expect(filteredByDate).toHaveLength(2);

    const filteredByItem = T.recoReportFilterTransactions(txns, {
      itemCodes: ["201050001"],
    });
    expect(filteredByItem).toHaveLength(1);
    expect(filteredByItem[0].itemCode).toBe("201050001");
  });
});

describe("Auto_Reco_Period_Report daily, weekly, and monthly aggregation", () => {
  const makeTxns = () =>
    T.recoReportParseRows({
      date: col([
        "02/11/2026",
        "01/11/2026",
        "02/11/2026",
        "09/11/2026",
        "01/12/2026",
        "",
      ]),
      type: col(["OUT", "IN", "IN", "OUT", "IN", "IN"]),
      itemCode: col([
        110702000516,
        110702000516,
        201050001,
        110702000516,
        201050001,
        110702000516,
      ]),
      qtyMeters: col([300, 100, 250.5, 400, 500, 999]),
    });

  it("aggregates daily summaries sorted chronologically and skips null dates", () => {
    const daily = T.recoReportAggregateByDay(makeTxns());
    expect(daily).toEqual([
      {
        periodKey: "2026-11-01",
        periodLabel: "01/11/2026",
        startDate: "2026-11-01",
        endDate: "2026-11-01",
        totalIn: 100,
        totalOut: 0,
        net: 100,
        inCount: 1,
        outCount: 0,
        totalCount: 1,
        uniqueItemCount: 1,
      },
      {
        periodKey: "2026-11-02",
        periodLabel: "02/11/2026",
        startDate: "2026-11-02",
        endDate: "2026-11-02",
        totalIn: 250.5,
        totalOut: 300,
        net: -49.5,
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        uniqueItemCount: 2,
      },
      {
        periodKey: "2026-11-09",
        periodLabel: "09/11/2026",
        startDate: "2026-11-09",
        endDate: "2026-11-09",
        totalIn: 0,
        totalOut: 400,
        net: -400,
        inCount: 0,
        outCount: 1,
        totalCount: 1,
        uniqueItemCount: 1,
      },
      {
        periodKey: "2026-12-01",
        periodLabel: "01/12/2026",
        startDate: "2026-12-01",
        endDate: "2026-12-01",
        totalIn: 500,
        totalOut: 0,
        net: 500,
        inCount: 1,
        outCount: 0,
        totalCount: 1,
        uniqueItemCount: 1,
      },
    ]);
  });

  it("computes ISO week descriptors accurately across year boundaries", () => {
    expect(T.recoReportGetIsoWeekDescriptor("2025-12-29")).toEqual({
      periodKey: "2026-W01",
      periodLabel: "Tuần 01/2026",
      startDate: "2025-12-29",
      endDate: "2026-01-04",
    });
    expect(T.recoReportGetIsoWeekDescriptor("2026-01-01")).toEqual({
      periodKey: "2026-W01",
      periodLabel: "Tuần 01/2026",
      startDate: "2025-12-29",
      endDate: "2026-01-04",
    });
    expect(T.recoReportGetIsoWeekDescriptor("2021-01-01")).toEqual({
      periodKey: "2020-W53",
      periodLabel: "Tuần 53/2020",
      startDate: "2020-12-28",
      endDate: "2021-01-03",
    });
  });

  it("aggregates weekly summaries by ISO week Monday–Sunday", () => {
    const weekly = T.recoReportAggregateByWeek(makeTxns());
    expect(weekly).toEqual([
      {
        periodKey: "2026-W44",
        periodLabel: "Tuần 44/2026",
        startDate: "2026-10-26",
        endDate: "2026-11-01",
        totalIn: 100,
        totalOut: 0,
        net: 100,
        inCount: 1,
        outCount: 0,
        totalCount: 1,
        uniqueItemCount: 1,
      },
      {
        periodKey: "2026-W45",
        periodLabel: "Tuần 45/2026",
        startDate: "2026-11-02",
        endDate: "2026-11-08",
        totalIn: 250.5,
        totalOut: 300,
        net: -49.5,
        inCount: 1,
        outCount: 1,
        totalCount: 2,
        uniqueItemCount: 2,
      },
      {
        periodKey: "2026-W46",
        periodLabel: "Tuần 46/2026",
        startDate: "2026-11-09",
        endDate: "2026-11-15",
        totalIn: 0,
        totalOut: 400,
        net: -400,
        inCount: 0,
        outCount: 1,
        totalCount: 1,
        uniqueItemCount: 1,
      },
      {
        periodKey: "2026-W49",
        periodLabel: "Tuần 49/2026",
        startDate: "2026-11-30",
        endDate: "2026-12-06",
        totalIn: 500,
        totalOut: 0,
        net: 500,
        inCount: 1,
        outCount: 0,
        totalCount: 1,
        uniqueItemCount: 1,
      },
    ]);
  });

  it("aggregates monthly summaries and handles leap-year February boundaries", () => {
    expect(T.recoReportGetMonthDescriptor("2024-02-15")).toEqual({
      periodKey: "2024-02",
      periodLabel: "02/2024",
      startDate: "2024-02-01",
      endDate: "2024-02-29",
    });
    expect(T.recoReportGetMonthDescriptor("2026-02-10")).toEqual({
      periodKey: "2026-02",
      periodLabel: "02/2026",
      startDate: "2026-02-01",
      endDate: "2026-02-28",
    });

    const monthly = T.recoReportAggregateByMonth(makeTxns());
    expect(monthly).toEqual([
      {
        periodKey: "2026-11",
        periodLabel: "11/2026",
        startDate: "2026-11-01",
        endDate: "2026-11-30",
        totalIn: 350.5,
        totalOut: 700,
        net: -349.5,
        inCount: 2,
        outCount: 2,
        totalCount: 4,
        uniqueItemCount: 2,
      },
      {
        periodKey: "2026-12",
        periodLabel: "12/2026",
        startDate: "2026-12-01",
        endDate: "2026-12-31",
        totalIn: 500,
        totalOut: 0,
        net: 500,
        inCount: 1,
        outCount: 0,
        totalCount: 1,
        uniqueItemCount: 1,
      },
    ]);
  });

  it("builds period sheet rows with Excel formulas, date serials, and TỔNG CỘNG row", () => {
    const monthly = T.recoReportAggregateByMonth(makeTxns());
    const rows = T.recoReportBuildPeriodRows(monthly, 6);
    expect(rows).toEqual([
      [
        1,
        "11/2026",
        T.recoReportIsoToSerial("2026-11-01"),
        T.recoReportIsoToSerial("2026-11-30"),
        350.5,
        700,
        "=E6-F6",
        2,
        2,
        "=H6+I6",
        2,
      ],
      [
        2,
        "12/2026",
        T.recoReportIsoToSerial("2026-12-01"),
        T.recoReportIsoToSerial("2026-12-31"),
        500,
        0,
        "=E7-F7",
        1,
        0,
        "=H7+I7",
        1,
      ],
      [
        "",
        "TỔNG CỘNG",
        "",
        "",
        "=SUM(E6:E7)",
        "=SUM(F6:F7)",
        "=SUM(G6:G7)",
        "=SUM(H6:H7)",
        "=SUM(I6:I7)",
        "=SUM(J6:J7)",
        "",
      ],
    ]);

    expect(T.recoReportBuildPeriodRows([], 6)).toEqual([]);
  });

  it("detects last data row via getUsedRange(true) and returns 0 on empty sheet", () => {
    const emptySheet = { getUsedRange: vi.fn(() => undefined) };
    expect(T.recoReportGetLastDataRow(emptySheet)).toBe(0);
    expect(emptySheet.getUsedRange).toHaveBeenCalledWith(true);

    const dataSheet = {
      getUsedRange: vi.fn(() => ({
        getLastRow: () => ({ getRowIndex: () => 19 }),
      })),
    };
    expect(T.recoReportGetLastDataRow(dataSheet)).toBe(20);
    expect(dataSheet.getUsedRange).toHaveBeenCalledWith(true);
  });
});
