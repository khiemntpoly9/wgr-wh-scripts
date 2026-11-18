import { beforeAll, describe, expect, it, vi } from "vitest";

let T;
beforeAll(async () => {
  await import("../fabric/Auto_PMC_Offset_Adjust.ts");
  T = globalThis.__AUTO_PMC_OFFSET_ADJUST_TESTABLES__;
});

describe("Auto_PMC_Offset_Adjust config, parsing, and template", () => {
  it("exposes expected config and 8-column headers", () => {
    expect(T.PMC_OFFSET_CONFIG).toEqual({
      sheetName: "PMC_ADJUST",
      headerRow: 5,
      startDataRow: 6,
      templateInputRowCount: 15,
    });
    expect(T.PMC_OFFSET_HEADERS).toEqual([
      "STT",
      "Mã hàng hóa (Kho)",
      "Số lượng cấp của kho",
      "Mã hàng hóa từ ERP",
      "Số lượng cấp của PMC",
      "Số lượng PMC sau điều chỉnh",
      "Chênh lệch bù trừ",
      "Ghi chú",
    ]);
  });

  it("parses Vietnamese and US decimal/thousands numbers preserving 4+ decimals", () => {
    expect(T.pmcParseNumber("86,7226")).toBe(86.7226);
    expect(T.pmcParseNumber("64.860,26")).toBe(64860.26);
    expect(T.pmcParseNumber("64,860.26")).toBe(64860.26);
    expect(T.pmcParseNumber(66700)).toBe(66700);
    expect(T.pmcParseNumber("")).toBe(0);
  });

  it("normalizes item codes and converts safe integers for Excel cells", () => {
    expect(T.pmcNormalizeItemCode("1.10703E+11")).toBe("110703000000");
    expect(T.pmcNormalizeItemCode(110703000003)).toBe("110703000003");
    expect(T.pmcToCellItemCode("110703000003")).toBe(110703000003);
    expect(T.pmcToCellItemCode("FAB-001")).toBe("FAB-001");
  });

  it("builds empty template rows without a TỔNG CỘNG blocker row", () => {
    const rows = T.pmcBuildEmptyTemplateRows(15, 6);
    expect(rows).toHaveLength(15);
    expect(rows[0]).toEqual([1, "", "", "", "", "", "", ""]);
    expect(rows[14]).toEqual([15, "", "", "", "", "", "", ""]);
  });
});

describe("Auto_PMC_Offset_Adjust extraction and offset calculation", () => {
  it("handles the user example: 110703000003 Kho=66700, PMC=[86,7226, 64.860,26]", () => {
    const rawAE = [
      [1, 110703000003, 66700, 110703000003, "86,7226"],
      [2, "", "", "110703000003", "64.860,26"],
      ["", "TỔNG CỘNG", 66700, "", 64946.9826],
    ];

    const parsed = T.pmcExtractInputEntries(rawAE);
    expect(parsed.khoEntries).toEqual([{ itemCode: "110703000003", qty: 66700 }]);
    expect(parsed.erpEntries).toEqual([
      { itemCode: "110703000003", qty: 86.7226 },
      { itemCode: "110703000003", qty: 64860.26 },
    ]);

    const adjusted = T.pmcComputeOffsetAdjustments(
      parsed.khoEntries,
      parsed.erpEntries,
    );
    expect(adjusted).toEqual([
      {
        itemCode: "110703000003",
        originalQty: 86.7226,
        adjustedQty: 86.7226,
        offsetDiff: 0,
        note: "",
        isLastOfItem: false,
        matchedKho: true,
        isShortageWarning: false,
      },
      {
        itemCode: "110703000003",
        originalQty: 64860.26,
        adjustedQty: 66613.2774,
        offsetDiff: 1753.0174,
        note: "Đã bù trừ",
        isLastOfItem: true,
        matchedKho: true,
        isShortageWarning: false,
      },
    ]);

    const totalAdjusted = T.pmcRoundQty(
      adjusted.reduce((acc, r) => acc + r.adjustedQty, 0),
    );
    expect(totalAdjusted).toBe(66700);

    const sheetRows = T.pmcBuildCalculatedRows(parsed.khoEntries, adjusted, 6);
    expect(sheetRows).toEqual([
      [1, 110703000003, 66700, 110703000003, 86.7226, 86.7226, 0, ""],
      [2, "", "", 110703000003, 64860.26, 66613.2774, 1753.0174, "Đã bù trừ"],
      [
        "",
        "TỔNG CỘNG",
        "=SUM(C6:C7)",
        "",
        "=SUM(E6:E7)",
        "=SUM(F6:F7)",
        "=SUM(G6:G7)",
        "",
      ],
    ]);
  });

  it("does NOT adjust and flags shortage warning when Kho < PMC, plus exact match and unmatched ERP", () => {
    const khoEntries = [
      { itemCode: "A1", qty: 100 },
      { itemCode: "B2", qty: 50 },
      { itemCode: "D4", qty: 80 },
    ];
    const erpEntries = [
      { itemCode: "A1", qty: 40 },
      { itemCode: "B2", qty: 60 },
      { itemCode: "D4", qty: 50 },
      { itemCode: "A1", qty: 60 },
      { itemCode: "D4", qty: 35.5 },
      { itemCode: "C3", qty: 25 },
    ];

    const adjusted = T.pmcComputeOffsetAdjustments(khoEntries, erpEntries);
    expect(adjusted).toEqual([
      {
        itemCode: "A1",
        originalQty: 40,
        adjustedQty: 40,
        offsetDiff: 0,
        note: "",
        isLastOfItem: false,
        matchedKho: true,
        isShortageWarning: false,
      },
      {
        itemCode: "B2",
        originalQty: 60,
        adjustedQty: 60,
        offsetDiff: 0,
        note: "Cảnh báo: Kho ít hơn PMC (thiếu 10) - Không bù",
        isLastOfItem: true,
        matchedKho: true,
        isShortageWarning: true,
      },
      {
        itemCode: "D4",
        originalQty: 50,
        adjustedQty: 50,
        offsetDiff: 0,
        note: "",
        isLastOfItem: false,
        matchedKho: true,
        isShortageWarning: true,
      },
      {
        itemCode: "A1",
        originalQty: 60,
        adjustedQty: 60,
        offsetDiff: 0,
        note: "Khớp",
        isLastOfItem: true,
        matchedKho: true,
        isShortageWarning: false,
      },
      {
        itemCode: "D4",
        originalQty: 35.5,
        adjustedQty: 35.5,
        offsetDiff: 0,
        note: "Cảnh báo: Kho ít hơn PMC (thiếu 5.5) - Không bù",
        isLastOfItem: true,
        matchedKho: true,
        isShortageWarning: true,
      },
      {
        itemCode: "C3",
        originalQty: 25,
        adjustedQty: 25,
        offsetDiff: 0,
        note: "Không tìm thấy mã bên Kho",
        isLastOfItem: true,
        matchedKho: false,
        isShortageWarning: false,
      },
    ]);
  });

  it("detects last data row via getUsedRange(true)", () => {
    const emptySheet = { getUsedRange: vi.fn(() => undefined) };
    expect(T.pmcGetLastDataRow(emptySheet)).toBe(0);
    expect(emptySheet.getUsedRange).toHaveBeenCalledWith(true);

    const dataSheet = {
      getUsedRange: vi.fn(() => ({
        getLastRow: () => ({ getRowIndex: () => 14 }),
      })),
    };
    expect(T.pmcGetLastDataRow(dataSheet)).toBe(15);
  });
});
