import { beforeAll, describe, expect, it } from "vitest";

import { createMockExcelScriptEnums } from "./helpers/excelMocks.js";

const getTestables = () => {
  const testables = globalThis.__AUTO_DAILY_REPORT_TESTABLES__;
  if (!testables) throw new Error("Auto_Daily_Report test helpers were not loaded.");
  return testables;
};

beforeAll(async () => {
  globalThis.ExcelScript = createMockExcelScriptEnums();
  await import("../Auto_Daily_Report.ts");
});

const CELLS = {
  outQty: { daily: "T9", month: "T10", decimals: 0 },
  outVolume: { daily: "N9", month: "N10", decimals: 2 },
  inCartons: { daily: "H9", month: "H10", decimals: 0 },
  inVolume: { daily: "B9", month: "B10", decimals: 2 },
};

/** Sheet ngày 29 vừa copy từ sheet ngày 28 (ô ngày đang là số của ngày 28). */
const copiedSheet = () => ({
  T9: 2790,
  T10: "=T9+1500+2000",
  T11: "=3+5+4",
  N9: 530.4599999,
  N10: "=N9+100.5+220.25",
  H9: 1250,
  H10: "=H9+800",
  B9: 85.3199,
  B10: "=B9+40.5",
});

const stateOf28 = () => ({
  date: "2026-09-28",
  prevDate: "2026-09-26",
  prevDaily: { outQty: 2000, outVolume: 220.25, inCartons: 800, inVolume: 40.5 },
  baseMonth: {
    outQty: "=T9+1500",
    outVolume: "=N9+100.5",
    inCartons: "=H9",
    inVolume: "=B9",
  },
  baseCount: "=3+5",
});

const defaultTerms = () => ({
  outQty: [1307, 1483],
  outVolume: [198.64, 331.82],
  inCartons: [900],
  inVolume: [60.25],
});

const makeInput = ({
  reportDate = "2026-09-29",
  sheet = copiedSheet(),
  terms = defaultTerms(),
  state = stateOf28(),
  initialPrevDate = null,
} = {}) => ({
  reportDate,
  state,
  initialPrevDate,
  metrics: Object.keys(CELLS).map((key) => ({
    key,
    dailyCell: CELLS[key].daily,
    decimals: CELLS[key].decimals,
    terms: terms[key],
    currentDailyValue: sheet[CELLS[key].daily],
    currentMonthFormula: sheet[CELLS[key].month],
  })),
  currentCountFormula: sheet.T11,
  containerCount: terms.outQty.length,
});

/** Ghi kết quả lên sheet như main. */
const applyResult = (sheet, result, values) => {
  const next = { ...sheet, T11: result.count };
  Object.keys(CELLS).forEach((key) => {
    next[CELLS[key].daily] = values[key];
    next[CELLS[key].month] = result.metrics[key].month;
  });
  return next;
};

describe("Auto_Daily_Report helpers", () => {
  it("đọc ngày A1 dạng Excel serial, text dd/mm/yyyy và yyyy-mm-dd", () => {
    const { dailyReportParseDate } = getTestables();
    expect(dailyReportParseDate(46293)).toBe("2026-09-28");
    expect(dailyReportParseDate("28/09/2026")).toBe("2026-09-28");
    expect(dailyReportParseDate("Ngày 1/9/2026")).toBe("2026-09-01");
    expect(dailyReportParseDate("2026-09-28")).toBe("2026-09-28");
    expect(dailyReportParseDate("31/02/2026")).toBeNull();
    expect(dailyReportParseDate("")).toBeNull();
  });

  it("đọc tên sheet dạng DDMM", () => {
    const { dailyReportParseSheetName } = getTestables();
    expect(dailyReportParseSheetName("2909")).toEqual({ day: 29, month: 9 });
    expect(dailyReportParseSheetName("0110")).toEqual({ day: 1, month: 10 });
    expect(dailyReportParseSheetName("2909 (2)")).toEqual({ day: 29, month: 9 });
    expect(dailyReportParseSheetName("29092026")).toBeNull();
    expect(dailyReportParseSheetName("3213")).toBeNull();
    expect(dailyReportParseSheetName("Sheet1")).toBeNull();
  });

  it("so sánh tên sheet với ngày A1", () => {
    const { dailyReportSheetNameMatches } = getTestables();
    expect(dailyReportSheetNameMatches("2026-09-29", { day: 29, month: 9 })).toBe(true);
    expect(dailyReportSheetNameMatches("2026-10-01", { day: 1, month: 10 })).toBe(true);
    expect(dailyReportSheetNameMatches("2026-09-28", { day: 29, month: 9 })).toBe(false);
  });

  it("suy ra năm từ ngày tham chiếu khi lấy ngày theo tên sheet", () => {
    const { dailyReportDateFromSheetName, dailyReportIsoToSerial, dailyReportParseDate } =
      getTestables();
    expect(dailyReportDateFromSheetName({ day: 1, month: 10 }, "2026-09-30")).toBe("2026-10-01");
    expect(dailyReportDateFromSheetName({ day: 2, month: 1 }, "2026-12-31")).toBe("2027-01-02");
    expect(dailyReportDateFromSheetName({ day: 31, month: 9 }, "2026-09-30")).toBeNull();
    expect(dailyReportParseDate(dailyReportIsoToSerial("2026-09-29"))).toBe("2026-09-29");
  });

  it("gom số lượng và thể tích xuất theo cont, dùng số phiếu khi thiếu cont, bỏ cont 0", () => {
    const { dailyReportGroupContainers } = getTestables();
    expect(
      dailyReportGroupContainers(
        [
          { dispatch_no: "PXK1", container_no: "abcu123", total_quantity: 1000, total_volume: 150.1234 },
          { dispatch_no: "PXK2", container_no: "ABCU123 ", total_quantity: 307, total_volume: "48.52" },
          { dispatch_no: "PXK3", container_no: "", total_quantity: 1483, total_volume: 331.8199 },
          { dispatch_no: "PXK4", container_no: "ZZZU9", total_quantity: 0, total_volume: 12 },
        ],
        2,
      ),
    ).toEqual({ quantities: [1307, 1483], volumes: [198.64, 331.82] });
  });

  it("lấy số thùng và thể tích nhập từ tổng của ngày", () => {
    const { dailyReportReceiptTerms } = getTestables();
    expect(
      dailyReportReceiptTerms({ carton_quantity: 1250, total_volume: 85.3249 }, 2),
    ).toEqual({ cartons: [1250], volumes: [85.32] });
    expect(
      dailyReportReceiptTerms({ carton_quantity: 0, total_volume: 0 }, 2),
    ).toEqual({ cartons: [], volumes: [] });
    expect(dailyReportReceiptTerms(undefined, 2)).toEqual({ cartons: [], volumes: [] });
  });

  it("đọc số có dấu phẩy hoặc dấu chấm thập phân", () => {
    const { dailyReportNumber } = getTestables();
    expect(dailyReportNumber(198.64)).toBe(198.64);
    expect(dailyReportNumber("198,64")).toBe(198.64);
    expect(dailyReportNumber("198.64")).toBe(198.64);
    expect(dailyReportNumber("")).toBe(0);
    expect(dailyReportNumber("abc")).toBe(0);
  });

  it("dựng và nối công thức cộng", () => {
    const { dailyReportBuildSumFormula, dailyReportAppendTerm } = getTestables();
    expect(dailyReportBuildSumFormula([1307, 1483])).toBe("=1307+1483");
    expect(dailyReportBuildSumFormula([])).toBe("=0");
    expect(dailyReportAppendTerm("=3+5", 2)).toBe("=3+5+2");
    expect(dailyReportAppendTerm("", 2)).toBe("=2");
    expect(dailyReportAppendTerm("=0", 2)).toBe("=2");
    expect(dailyReportAppendTerm(8, 2)).toBe("=8+2");
  });

  it("đọc và ghi trạng thái, bỏ qua nội dung không hợp lệ", () => {
    const { dailyReportParseState, dailyReportSerializeState } = getTestables();
    const state = stateOf28();
    expect(dailyReportParseState(dailyReportSerializeState(state))).toEqual(state);
    expect(dailyReportParseState("")).toBeNull();
    expect(dailyReportParseState("{not json")).toBeNull();
    expect(dailyReportParseState('{"date":"x"}')).toBeNull();
  });

  it("đọc trạng thái định dạng cũ (field phẳng T9/T10/T11/N9/N10)", () => {
    const { dailyReportParseState } = getTestables();
    expect(
      dailyReportParseState(
        JSON.stringify({
          date: "2026-09-29",
          prevDate: "2026-09-28",
          prevT9Value: 2790,
          baseT10: "=T9+1500",
          baseT11: "=3+5+4",
          prevN9Value: 530.46,
          baseN10: "=N9+100.5",
        }),
      ),
    ).toEqual({
      date: "2026-09-29",
      prevDate: "2026-09-28",
      prevDaily: { outQty: 2790, outVolume: 530.46 },
      baseMonth: { outQty: "=T9+1500", outVolume: "=N9+100.5" },
      baseCount: "=3+5+4",
    });
  });
});

describe("bảng tồn kho quy cont", () => {
  it("liệt kê cột từ B đến U", () => {
    const { dailyReportColumnRange } = getTestables();
    const columns = dailyReportColumnRange("B", "U");
    expect(columns).toHaveLength(20);
    expect(columns[0]).toBe("B");
    expect(columns[19]).toBe("U");
    expect(dailyReportColumnRange("Y", "AB")).toEqual(["Y", "Z", "AA", "AB"]);
  });

  it("gom thể tích xuất / nhập theo mã thị trường + phân loại", () => {
    const { dailyReportDispatchVolumeByCode, dailyReportReceiptVolumeByCode } =
      getTestables();
    expect(
      dailyReportDispatchVolumeByCode([
        { code: "euf", total_volume: 10.5 },
        { code: "EUS", total_volume: "3.25" },
        { code: "", total_volume: 99 },
      ]),
    ).toEqual({ EUF: 10.5, EUS: 3.25 });
    expect(
      dailyReportReceiptVolumeByCode([
        {
          market_code: "EU",
          classifications: [
            { classification_code: "F", total_volume: 20 },
            { classification_code: "S", total_volume: 5.5 },
          ],
        },
        { market_code: "UE", classifications: [{ classification_code: "T", total_volume: 7 }] },
      ]),
    ).toEqual({ EUF: 20, EUS: 5.5, UET: 7 });
  });

  it("dựng công thức xuất / nhập chia 68 và tồn cuối =đầu-xuất+nhập", () => {
    const { dailyReportStockRows } = getTestables();
    const result = dailyReportStockRows({
      codes: ["EUF", "EUS", "", "UET"],
      columns: ["B", "C", "D", "E"],
      openRow: 4,
      outRow: 5,
      inRow: 6,
      outByCode: { EUF: 198.6449, UET: 0.001, AUF: 12 },
      inByCode: { EUS: 60.25, XXX: 0 },
      containerCbm: 68,
      volumeDecimals: 2,
    });
    expect(result.outFormulas).toEqual(["=198.64/68", "", "", ""]);
    expect(result.inFormulas).toEqual(["", "=60.25/68", "", ""]);
    expect(result.closeFormulas).toEqual(["=B4-B5+B6", "=C4-C5+C6", "", "=E4-E5+E6"]);
    expect(result.unmatchedOut).toEqual(["AUF"]);
    expect(result.unmatchedIn).toEqual([]);
  });

  it("trạng thái giữ cờ đã copy tồn đầu ngày", () => {
    const { dailyReportParseState, dailyReportSerializeState } = getTestables();
    const state = { ...stateOf28(), stockOpened: true };
    expect(dailyReportParseState(dailyReportSerializeState(state))).toEqual(state);
  });
});

describe("dailyReportCompute", () => {
  it("ngày thường: ô ngày theo dữ liệu mới, lũy kế tháng nối tổng hôm trước", () => {
    const { dailyReportCompute } = getTestables();
    const result = dailyReportCompute(makeInput());
    expect(result.monthReset).toBe(false);
    expect(result.metrics.outQty).toMatchObject({
      daily: "=1307+1483",
      month: "=T9+1500+2000+2790",
    });
    expect(result.metrics.outVolume).toMatchObject({
      daily: "=198.64+331.82",
      month: "=N9+100.5+220.25+530.46",
    });
    expect(result.metrics.inCartons).toMatchObject({
      daily: "=900",
      month: "=H9+800+1250",
    });
    expect(result.metrics.inVolume).toMatchObject({
      daily: "=60.25",
      month: "=B9+40.5+85.32",
    });
    expect(result.count).toBe("=3+5+4+2");
    expect(result.state).toEqual({
      date: "2026-09-29",
      prevDate: "2026-09-28",
      prevDaily: { outQty: 2790, outVolume: 530.46, inCartons: 1250, inVolume: 85.32 },
      baseMonth: {
        outQty: "=T9+1500+2000",
        outVolume: "=N9+100.5+220.25",
        inCartons: "=H9+800",
        inVolume: "=B9+40.5",
      },
      baseCount: "=3+5+4",
    });
  });

  it("ngày không xuất, không nhập: ô ngày =0, T11 giữ nguyên", () => {
    const { dailyReportCompute } = getTestables();
    const result = dailyReportCompute(
      makeInput({
        terms: { outQty: [], outVolume: [], inCartons: [], inVolume: [] },
      }),
    );
    expect(result.metrics.outQty.daily).toBe("=0");
    expect(result.metrics.outVolume.daily).toBe("=0");
    expect(result.metrics.inCartons.daily).toBe("=0");
    expect(result.metrics.inVolume.daily).toBe("=0");
    expect(result.metrics.outQty.month).toBe("=T9+1500+2000+2790");
    expect(result.count).toBe("=3+5+4");
  });

  it("hôm trước bằng 0 thì không nối vào lũy kế tháng", () => {
    const { dailyReportCompute } = getTestables();
    const result = dailyReportCompute(
      makeInput({ sheet: { ...copiedSheet(), T9: 0, H9: 0 } }),
    );
    expect(result.metrics.outQty.month).toBe("=T9+1500+2000");
    expect(result.metrics.outQty.appendedPrev).toBe(0);
    expect(result.metrics.inCartons.month).toBe("=H9+800");
  });

  it("chạy lại cùng ngày không cộng trùng và cập nhật số liệu mới", () => {
    const { dailyReportCompute } = getTestables();
    const first = dailyReportCompute(makeInput());
    const sheetAfterFirst = applyResult(copiedSheet(), first, {
      outQty: 2790,
      outVolume: 530.46,
      inCartons: 900,
      inVolume: 60.25,
    });
    const rerun = dailyReportCompute(
      makeInput({
        sheet: sheetAfterFirst,
        state: first.state,
        terms: {
          outQty: [1307, 1483, 900],
          outVolume: [198.64, 331.82, 60.5],
          inCartons: [1000],
          inVolume: [70],
        },
      }),
    );
    expect(rerun.metrics.outQty.daily).toBe("=1307+1483+900");
    expect(rerun.metrics.outQty.month).toBe("=T9+1500+2000+2790");
    expect(rerun.metrics.outVolume.month).toBe("=N9+100.5+220.25+530.46");
    expect(rerun.metrics.inCartons.daily).toBe("=1000");
    expect(rerun.metrics.inCartons.month).toBe("=H9+800+1250");
    expect(rerun.metrics.inVolume.month).toBe("=B9+40.5+85.32");
    expect(rerun.count).toBe("=3+5+4+3");
    expect(rerun.state).toEqual(first.state);
  });

  it("chạy lại với trạng thái cũ chưa có số nhập: lấy H/B hiện tại trên sheet", () => {
    const { dailyReportCompute } = getTestables();
    const result = dailyReportCompute(
      makeInput({
        state: {
          date: "2026-09-29",
          prevDate: "2026-09-28",
          prevDaily: { outQty: 2790, outVolume: 530.46 },
          baseMonth: { outQty: "=T9+1500+2000", outVolume: "=N9+100.5+220.25" },
          baseCount: "=3+5+4",
        },
      }),
    );
    expect(result.metrics.outQty.month).toBe("=T9+1500+2000+2790");
    expect(result.metrics.inCartons.month).toBe("=H9+800+1250");
    expect(result.metrics.inVolume.month).toBe("=B9+40.5+85.32");
  });

  it("sang tháng mới: lũy kế tháng = ô ngày, T11 bắt đầu lại", () => {
    const { dailyReportCompute } = getTestables();
    const result = dailyReportCompute(
      makeInput({
        reportDate: "2026-10-01",
        state: { ...stateOf28(), date: "2026-09-30" },
      }),
    );
    expect(result.monthReset).toBe(true);
    expect(result.metrics.outQty.month).toBe("=T9");
    expect(result.metrics.outVolume.month).toBe("=N9");
    expect(result.metrics.inCartons.month).toBe("=H9");
    expect(result.metrics.inVolume.month).toBe("=B9");
    expect(result.count).toBe("=2");

    const empty = dailyReportCompute(
      makeInput({
        reportDate: "2026-10-01",
        state: { ...stateOf28(), date: "2026-09-30" },
        terms: { outQty: [], outVolume: [], inCartons: [], inVolume: [] },
      }),
    );
    expect(empty.count).toBe("=0");
  });

  it("lần chạy đầu không có trạng thái: dùng initialPrevDate nếu có", () => {
    const { dailyReportCompute } = getTestables();
    const sameMonth = dailyReportCompute(makeInput({ state: null }));
    expect(sameMonth.metrics.outQty.month).toBe("=T9+1500+2000+2790");
    expect(sameMonth.state.prevDate).toBeNull();

    const newMonth = dailyReportCompute(
      makeInput({
        reportDate: "2026-10-01",
        state: null,
        initialPrevDate: "2026-09-30",
      }),
    );
    expect(newMonth.metrics.outQty.month).toBe("=T9");
    expect(newMonth.metrics.inCartons.month).toBe("=H9");
    expect(newMonth.count).toBe("=2");
  });

  it("ô lũy kế chưa có công thức thì bắt đầu từ ô ngày", () => {
    const { dailyReportCompute } = getTestables();
    const result = dailyReportCompute(
      makeInput({
        state: null,
        sheet: { ...copiedSheet(), T10: "", T9: 500, H10: null, H9: 40 },
      }),
    );
    expect(result.metrics.outQty.month).toBe("=T9+500");
    expect(result.metrics.inCartons.month).toBe("=H9+40");
  });
});

describe("dailyReportLocalAddress", () => {
  it("bỏ tên sheet và dấu $ khỏi địa chỉ của Name", () => {
    const { dailyReportLocalAddress } = getTestables();
    expect(dailyReportLocalAddress("2909!$V$9")).toBe("V9");
    expect(dailyReportLocalAddress("'Sheet 1'!B3:W3")).toBe("B3:W3");
    expect(dailyReportLocalAddress("T9")).toBe("T9");
  });
});
