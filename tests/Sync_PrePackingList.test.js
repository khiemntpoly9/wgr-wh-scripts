import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__SYNC_PRE_PACKING_TESTABLES__;

  if (!testables) {
    throw new Error("Sync_PrePackingList test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../finished-goods/Sync_PrePackingList.ts");
});

const HEADER = "STT\n序号";
const HEADER_INDEX = 5;

/** Dòng dữ liệu 18 cột A→R theo format Auto_PXK. */
const dataRow = (stt, contractNo, sku, overrides = {}) => {
  const row = [
    stt,
    contractNo,
    sku,
    `Hàng ${sku}`,
    "80*45*120",
    10,
    "PCS",
    0,
    10,
    "thùng",
    "L1",
    2.5,
    25,
    0.1,
    1,
    "C12",
    "",
    "",
  ];
  Object.entries(overrides).forEach(([col, value]) => {
    row[Number(col)] = value;
  });
  return row;
};

const emptyRow = () => Array(18).fill("");

/**
 * Dựng values/formulas cho sheet: 5 dòng meta, tiêu đề ở index 5, dữ liệu từ index 6, dòng tổng ngay sau.
 * M, O của dòng dữ liệu là công thức ROUND như Auto_PXK.
 */
const buildSheet = (rows, { totals = "formula", formulaColumns = true } = {}) => {
  const values = [];
  const formulas = [];
  for (let i = 0; i < HEADER_INDEX; i++) {
    values.push(emptyRow());
    formulas.push(emptyRow());
  }
  values[3][3] = "NO: 000631";

  const header = emptyRow();
  header[0] = HEADER;
  header[12] = "TỔNG TRỌNG LƯỢNG (G.W./KG)总毛重";
  values.push(header);
  formulas.push(emptyRow());

  rows.forEach((row, i) => {
    const sheetRow = HEADER_INDEX + 2 + i;
    values.push(row);
    const f = emptyRow();
    if (formulaColumns && row[3] !== "") {
      f[12] = `=ROUND(L${sheetRow}*I${sheetRow}, 2)`;
      f[14] = `=ROUND(N${sheetRow}*I${sheetRow}, 2)`;
    }
    formulas.push(f);
  });

  const first = HEADER_INDEX + 2;
  const last = HEADER_INDEX + 1 + rows.length;
  if (totals === "formula") {
    const t = emptyRow();
    const tf = emptyRow();
    t[8] = 100;
    tf[8] = `=SUM(I${first}:I${last})`;
    t[12] = 200;
    tf[12] = `=SUM(M${first}:M${last})`;
    values.push(t);
    formulas.push(tf);
  } else if (totals === "text") {
    const t = emptyRow();
    t[0] = "TỔNG CỘNG";
    values.push(t);
    formulas.push(emptyRow());
  }

  return { values, formulas };
};

/** Dòng hệ thống khớp đúng dataRow cùng tham số. */
const systemLine = (lineNo, contractNo, sku, overrides = {}) => ({
  line_no: lineNo,
  contract_no: contractNo,
  sku: String(sku),
  product_name: `Hàng ${sku}`,
  dimensions: "80*45*120",
  dispatch_quantity: 10,
  unit: "PCS",
  previously_shipped_quantity: 0,
  packed_quantity: 10,
  pack_unit: "thùng",
  batch_no: "L1",
  unit_weight: 2.5,
  total_weight: 25,
  unit_volume: 0.1,
  total_volume: 1,
  location: "C12",
  notes: "",
  pmc_approved: false,
  source: "excel",
  ...overrides,
});

const readSheet = (rows, options) => {
  const { pplSyncDocSheet } = getTestables();
  const { values, formulas } = buildSheet(rows, options);
  return pplSyncDocSheet(values, formulas, ["STT", "序号"]);
};

describe("pplSyncKey", () => {
  it("normalizes SKU suffixes and contract case", () => {
    const { pplSyncKey } = getTestables();

    expect(pplSyncKey("912345C", " euf6001-1a ")).toBe("912345|EUF6001-1A");
    expect(pplSyncKey(912345, "EUF6001-1A")).toBe("912345|EUF6001-1A");
    expect(pplSyncKey("912345A", "X")).not.toBe(pplSyncKey("912345B", "X"));
    expect(pplSyncKey("", "EUF6001-1A")).toBe("");
    expect(pplSyncKey("912345", "")).toBe("");
  });
});

describe("pplSyncDocSheet", () => {
  it("finds header, data region, totals row and formula columns", () => {
    const info = readSheet([
      dataRow(1, "EUF6001-1A", 911111),
      dataRow(2, "EUF6001-1A", 922222),
      dataRow(3, "EUF6002-1A", 933333),
    ]);

    expect(info.headerIndex).toBe(5);
    expect(info.startIndex).toBe(6);
    expect(info.lastIndex).toBe(8);
    expect(info.totalsIndex).toBe(9);
    expect(info.formulaCols).toEqual([12, 14]);
    expect(info.rows).toHaveLength(3);
    expect(info.rows[0].rowIndex).toBe(6);
    expect(info.rows[0].key).toBe("911111|EUF6001-1A");
  });

  it("detects totals row by TỔNG text and keeps rows without key", () => {
    const blank = emptyRow();
    blank[3] = "Ghi chú chen giữa";
    const info = readSheet([dataRow(1, "EUF6001-1A", 911111), blank], {
      totals: "text",
      formulaColumns: false,
    });

    expect(info.totalsIndex).toBe(8);
    expect(info.formulaCols).toEqual([]);
    expect(info.rows).toHaveLength(2);
    expect(info.rows[1].key).toBe("");
  });

  it("returns totalsIndex -1 when there is no totals row", () => {
    const info = readSheet([dataRow(1, "EUF6001-1A", 911111)], { totals: "none" });

    expect(info.totalsIndex).toBe(-1);
    expect(info.lastIndex).toBe(6);
  });

  it("returns null without header row", () => {
    const { pplSyncDocSheet } = getTestables();

    expect(pplSyncDocSheet([emptyRow()], [emptyRow()], ["STT", "序号"])).toBeNull();
  });
});

describe("pplSyncLapKeHoach – thêm dòng", () => {
  const baseRows = () => [
    dataRow(1, "EUF6001-1A", 911111),
    dataRow(2, "EUF6001-1A", 922222),
    dataRow(3, "EUF6002-1A", 933333),
  ];
  const baseLines = () => [
    systemLine(1, "EUF6001-1A", 911111),
    systemLine(2, "EUF6001-1A", 922222),
    systemLine(3, "EUF6002-1A", 933333),
  ];

  it("returns an empty plan when sheet already matches", () => {
    const { pplSyncLapKeHoach, pplSyncKeHoachRong } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet(baseRows()), baseLines());

    expect(plan.inserts).toEqual([]);
    expect(plan.updates).toEqual([]);
    expect(plan.deletes).toEqual([]);
    expect(pplSyncKeHoachRong(plan)).toBe(true);
  });

  it("inserts below the last row of the same contract", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet(baseRows()), [
      ...baseLines(),
      systemLine(4, "EUF6001-1A", 984945, { source: "system", location: "C73, C74" }),
    ]);

    expect(plan.inserts).toHaveLength(1);
    expect(plan.inserts[0].afterRowIndex).toBe(7);
    expect(plan.inserts[0].lines.map((l) => l.sku)).toEqual(["984945"]);
    expect(plan.finalRows.map((r) => r.values[2])).toEqual([911111, 922222, "984945", 933333]);
    expect(plan.finalRows.map((r) => r.values[0])).toEqual([1, 2, 3, 4]);
    expect(plan.finalRows[2].mark).toBe("added");
    expect(plan.finalRows[0].mark).toBeNull();
  });

  it("fills all 18 columns of a new row and uses formulas for formula columns", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet(baseRows()), [
      ...baseLines(),
      systemLine(4, "EUF6001-1A", 984945, {
        product_name: "Tủ gỗ 2 cánh",
        packed_quantity: 12,
        notes: "Dễ vỡ",
        pmc_approved: true,
        location: "C73, C74",
      }),
    ]);

    // Dòng mới đứng thứ 3 trong vùng dữ liệu → dòng Excel số 6 + 2 + 1 = 9.
    expect(plan.finalRows[2].values).toEqual([
      3,
      "EUF6001-1A",
      "984945",
      "Tủ gỗ 2 cánh",
      "80*45*120",
      10,
      "PCS",
      0,
      12,
      "thùng",
      "L1",
      2.5,
      "=ROUND(L9*I9, 2)",
      0.1,
      "=ROUND(N9*I9, 2)",
      "C73, C74",
      "Dễ vỡ",
      "Y",
    ]);
  });

  it("writes totals as values when the sheet has no formula columns", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const info = readSheet(baseRows(), { formulaColumns: false });
    const plan = pplSyncLapKeHoach(info, [
      ...baseLines(),
      systemLine(4, "EUF6009-1A", 984945, { total_weight: 306, total_volume: 5.184 }),
    ]);

    const added = plan.finalRows[3].values;
    expect(added[12]).toBe(306);
    expect(added[14]).toBe(5.184);
  });

  it("appends a new contract after the last data row", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet(baseRows()), [
      ...baseLines(),
      systemLine(4, "EUF6009-1A", 984945),
    ]);

    expect(plan.inserts).toEqual([
      { afterRowIndex: 8, lines: [expect.objectContaining({ sku: "984945" })] },
    ]);
  });

  it("keeps system order for several new rows sharing an insertion point", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet(baseRows()), [
      ...baseLines(),
      systemLine(5, "EUF6001-1A", 955555),
      systemLine(4, "EUF6001-1A", 944444),
    ]);

    expect(plan.inserts).toHaveLength(1);
    expect(plan.inserts[0].lines.map((l) => l.sku)).toEqual(["955555", "944444"]);
  });

  it("matches duplicate keys by occurrence order", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const rows = [dataRow(1, "EUF6001-1A", 911111), dataRow(2, "EUF6001-1A", 911111)];
    const lines = [
      systemLine(1, "EUF6001-1A", 911111),
      systemLine(2, "EUF6001-1A", 911111),
      systemLine(3, "EUF6001-1A", 911111),
    ];

    const plan = pplSyncLapKeHoach(readSheet(rows), lines);

    expect(plan.inserts).toHaveLength(1);
    expect(plan.inserts[0].afterRowIndex).toBe(7);
  });

  it("inserts after the header when the sheet has no data rows", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([]), [systemLine(1, "EUF6001-1A", 911111)]);

    expect(plan.inserts[0].afterRowIndex).toBe(HEADER_INDEX);
    expect(plan.finalRows).toHaveLength(1);
  });

  it("does not touch rows without SKU or contract", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const note = emptyRow();
    note[0] = "x";
    note[3] = "Ghi chú";
    const plan = pplSyncLapKeHoach(
      readSheet([dataRow(1, "EUF6001-1A", 911111), note, dataRow(2, "EUF6001-1A", 922222)]),
      [systemLine(1, "EUF6001-1A", 911111), systemLine(2, "EUF6001-1A", 922222)],
    );

    expect(plan.deletes).toEqual([]);
    expect(plan.finalRows.map((r) => r.values[0])).toEqual([1, "x", 2]);
  });
});

describe("pplSyncLapKeHoach – cập nhật và xoá", () => {
  const columnsOf = (plan) => plan.updates.flatMap((u) => u.changes.map((c) => c.column));

  it("reports a change for every compared column that differs", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(
      readSheet([dataRow(1, "EUF6001-1A", 911111, { 16: "ghi chú cũ" })], { formulaColumns: false }),
      [
        systemLine(1, "EUF6001-1A", 911111, {
          dispatch_quantity: 12,
          previously_shipped_quantity: 2,
          packed_quantity: 15,
          batch_no: "L2608",
          unit_weight: 2.6,
          total_weight: 39,
          unit_volume: 0.2,
          total_volume: 3,
          location: "C73, C74",
          notes: "Dễ vỡ",
          pmc_approved: true,
        }),
      ],
    );

    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0].rowIndex).toBe(6);
    expect(columnsOf(plan)).toEqual(["F", "H", "I", "K", "L", "M", "N", "O", "P", "Q", "R"]);
    expect(plan.finalRows[0].mark).toBe("updated");
    expect(plan.finalRows[0].values.slice(5)).toEqual([
      12, "PCS", 2, 15, "thùng", "L2608", 2.6, 39, 0.2, 3, "C73, C74", "Dễ vỡ", "Y",
    ]);
  });

  it("ignores differences within tolerance", () => {
    const { pplSyncLapKeHoach, pplSyncKeHoachRong } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6001-1A", 911111)], { formulaColumns: false }), [
      systemLine(1, "EUF6001-1A", 911111, {
        unit_weight: 2.5004,
        total_weight: 25.004,
        unit_volume: 0.10004,
        total_volume: 1.004,
      }),
    ]);

    expect(pplSyncKeHoachRong(plan)).toBe(true);
  });

  it("does not compare M and O when they are formulas", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6001-1A", 911111)]), [
      systemLine(1, "EUF6001-1A", 911111, { total_weight: 99, total_volume: 9 }),
    ]);

    expect(plan.updates).toEqual([]);
  });

  it("rewrites formulas of an updated row at its new position", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(
      readSheet([dataRow(1, "EUF6001-1A", 911111), dataRow(2, "EUF6001-1A", 922222)]),
      [systemLine(1, "EUF6001-1A", 911111), systemLine(2, "EUF6001-1A", 922222, { packed_quantity: 15 })],
    );

    expect(plan.finalRows[1].values[8]).toBe(15);
    expect(plan.finalRows[1].values[12]).toBe("=ROUND(L8*I8, 2)");
    expect(plan.finalRows[1].values[14]).toBe("=ROUND(N8*I8, 2)");
  });

  it("keeps the Excel batch number when the system has none", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6001-1A", 911111, { 10: "2608150001" })]), [
      systemLine(1, "EUF6001-1A", 911111, { batch_no: "" }),
    ]);

    expect(plan.updates).toEqual([]);
  });

  it("only overwrites notes when system has notes and PMC when approved", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(
      readSheet([dataRow(1, "EUF6001-1A", 911111, { 16: "Ghi chú tay", 17: "Y" })]),
      [systemLine(1, "EUF6001-1A", 911111, { notes: "", pmc_approved: false })],
    );

    expect(plan.updates).toEqual([]);
  });

  it("never changes columns B, C, D, E, G, J of existing rows", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(
      readSheet([dataRow(1, "euf6001-1a", "911111C", { 3: "Tên Excel", 4: "1*2*3", 6: "BỘ", 9: "kiện" })]),
      [
        systemLine(1, "EUF6001-1A", 911111, {
          product_name: "Tên hệ thống",
          dimensions: "9*9*9",
          unit: "PCS",
          pack_unit: "thùng",
          packed_quantity: 11,
        }),
      ],
    );

    expect(columnsOf(plan)).toEqual(["I"]);
    expect(plan.finalRows[0].values.slice(1, 5)).toEqual(["euf6001-1a", "911111C", "Tên Excel", "1*2*3"]);
    expect(plan.finalRows[0].values[6]).toBe("BỘ");
    expect(plan.finalRows[0].values[9]).toBe("kiện");
  });

  it("overwrites manual Excel edits and records old and new values", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6001-1A", 911111, { 8: 99 })]), [
      systemLine(1, "EUF6001-1A", 911111),
    ]);

    expect(plan.updates[0].changes).toEqual([
      {
        sku: "911111",
        contractNo: "EUF6001-1A",
        column: "I",
        label: "SL đóng gói",
        oldValue: 99,
        newValue: 10,
      },
    ]);
  });

  it("deletes keyed rows missing on the system and keeps rows without key", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const note = emptyRow();
    note[3] = "Ghi chú";
    const plan = pplSyncLapKeHoach(
      readSheet([dataRow(1, "EUF6001-1A", 911111), dataRow(2, "EUF6001-1A", 922222), note]),
      [systemLine(1, "EUF6001-1A", 911111)],
    );

    expect(plan.deletes).toEqual([7]);
    expect(plan.finalRows.map((r) => r.rowIndex)).toEqual([6, 8]);
    expect(plan.summary.deleted).toEqual([{ sku: "922222", contractNo: "EUF6001-1A" }]);
  });

  it("computes insertion points without deleted rows", () => {
    const { pplSyncLapKeHoach } = getTestables();
    const plan = pplSyncLapKeHoach(
      readSheet([
        dataRow(1, "EUF6001-1A", 911111),
        dataRow(2, "EUF6001-1A", 922222),
        dataRow(3, "EUF6002-1A", 933333),
      ]),
      [
        systemLine(1, "EUF6001-1A", 911111),
        systemLine(2, "EUF6001-1A", 944444),
        systemLine(3, "EUF6009-1A", 955555),
      ],
    );

    expect(plan.deletes).toEqual([7, 8]);
    expect(plan.inserts).toEqual([
      {
        afterRowIndex: 6,
        lines: [expect.objectContaining({ sku: "944444" }), expect.objectContaining({ sku: "955555" })],
      },
    ]);
    expect(plan.finalRows.map((r) => r.values[0])).toEqual([1, 2, 3]);
  });
});

describe("pplSyncTomTat – cập nhật và xoá", () => {
  it("lists deleted rows and overwritten cells", () => {
    const { pplSyncLapKeHoach, pplSyncTomTat } = getTestables();
    const plan = pplSyncLapKeHoach(
      readSheet([
        dataRow(1, "EUF6001-1A", 911111),
        dataRow(2, "EUF6001-1A", 922222),
        dataRow(3, "EUF6001-1A", 933333, { 10: "" }),
      ]),
      [
        systemLine(1, "EUF6001-1A", 911111, { packed_quantity: 15 }),
        systemLine(3, "EUF6001-1A", 933333, { batch_no: "L2608" }),
      ],
    );

    const text = pplSyncTomTat(plan, { listNo: "000631", status: "active" }, 2, 50);

    expect(text).toContain("Thêm 0 · Cập nhật 2 · Xoá 1");
    expect(text).toContain("Đã xoá:\n  - 922222 · EUF6001-1A");
    expect(text).toContain("Ô đã ghi đè theo hệ thống (nền vàng):");
    expect(text).toContain("  ~ 911111 · EUF6001-1A · SL đóng gói (I): 10 → 15");
    expect(text).toContain("  ~ 933333 · EUF6001-1A · Số lô (K): (trống) → L2608");
  });
});

describe("pplSyncDocBangSoLo", () => {
  it("maps trimmed upper-case contracts to lots and skips the header row", () => {
    const { pplSyncDocBangSoLo } = getTestables();
    const map = pplSyncDocBangSoLo([
      ["HỢP ĐỒNG", "SỐ LÔ"],
      [" euf6001-1a ", 2608150001],
      ["EUF6002-1A", ""],
      ["", "2608150009"],
    ]);

    expect(Array.from(map.entries())).toEqual([["EUF6001-1A", "2608150001"]]);
  });
});

describe("pplSyncDienDongMoi", () => {
  const UNIT_CFG = {
    qtyUnitDefault: "PCS",
    packUnitCarton: "箱/ THÙNG CARTON",
    packUnitPackage: "PACKAGE",
    packageHeight: 1.1,
  };
  const planWith = (rows, lines) => getTestables().pplSyncLapKeHoach(readSheet(rows), lines);

  it("sets G to PCS and J by package height on added rows only", () => {
    const { pplSyncDienDongMoi } = getTestables();
    const plan = planWith(
      [dataRow(1, "EUF6001-1A", 911111, { 6: "BỘ", 9: "kiện" })],
      [
        systemLine(1, "EUF6001-1A", 911111),
        systemLine(2, "EUF6001-1A", 922222, { unit: "Cái", pack_unit: "" }),
        systemLine(3, "EUF6001-1A", 933333, { dimensions: "44*24*1,1", pack_unit: "" }),
      ],
    );

    pplSyncDienDongMoi(plan, new Map(), UNIT_CFG);

    expect(plan.finalRows.map((r) => [r.values[6], r.values[9]])).toEqual([
      ["BỘ", "kiện"],
      ["PCS", "箱/ THÙNG CARTON"],
      ["PCS", "PACKAGE"],
    ]);
  });

  it("keeps a system batch number on added rows", () => {
    const { pplSyncDienDongMoi } = getTestables();
    const plan = planWith([], [systemLine(1, "EUF6001-1A", 911111, { batch_no: "2608150005" })]);

    expect(pplSyncDienDongMoi(plan, new Map([["EUF6001-1A", "2608150001"]]), UNIT_CFG)).toEqual([]);
    expect(plan.finalRows[0].values[10]).toBe("2608150005");
  });

  it("fills a missing batch number from the lot-no sheet first", () => {
    const { pplSyncDienDongMoi } = getTestables();
    const plan = planWith(
      [dataRow(1, "EUF6001-1A", 911111, { 10: "2608150002" })],
      [systemLine(1, "EUF6001-1A", 911111, { batch_no: "2608150002" }), systemLine(2, "euf6001-1a", 922222, { batch_no: "" })],
    );

    pplSyncDienDongMoi(plan, new Map([["EUF6001-1A", "2608150001"]]), UNIT_CFG);

    expect(plan.finalRows[1].values[10]).toBe("2608150001");
  });

  it("falls back to a 10-digit lot of a row sharing the base contract", () => {
    const { pplSyncDienDongMoi } = getTestables();
    const plan = planWith(
      [dataRow(1, "EUF6001-1A", 911111, { 10: "2608150002" })],
      [
        systemLine(1, "EUF6001-1A", 911111, { batch_no: "2608150002" }),
        systemLine(2, "EUF6001 3", 922222, { batch_no: "" }),
      ],
    );

    expect(pplSyncDienDongMoi(plan, new Map(), UNIT_CFG)).toEqual([]);
    expect(plan.finalRows[1].values[10]).toBe("2608150002");
  });

  it("returns positions of added rows still missing a lot and lists them in the summary", () => {
    const { pplSyncDienDongMoi, pplSyncTomTat } = getTestables();
    const plan = planWith(
      [dataRow(1, "EUF6001-1A", 911111)],
      [systemLine(1, "EUF6001-1A", 911111), systemLine(2, "EUF6009-1A", 984945, { batch_no: "0" })],
    );

    expect(pplSyncDienDongMoi(plan, new Map(), UNIT_CFG)).toEqual([1]);
    expect(plan.finalRows[1].values[10]).toBe("");
    expect(plan.summary.lotMissing).toEqual([{ sku: "984945", contractNo: "EUF6009-1A" }]);
    expect(pplSyncTomTat(plan, { listNo: "000631", status: "active" }, 1, 50)).toContain(
      "Dòng bổ sung chưa có số lô (ô K tô đỏ):\n  ! 984945 · EUF6009-1A",
    );
  });
});

describe("pplSyncDongTong", () => {
  it("rewrites SUM range to the new data bounds", () => {
    const { pplSyncDongTong } = getTestables();

    expect(pplSyncDongTong("=SUM(M5:M85)", 5, 86)).toBe("=SUM(M5:M86)");
    expect(pplSyncDongTong("=sum($I$7:$I$9)", 7, 8)).toBe("=SUM(I7:I8)");
    expect(pplSyncDongTong("=SUM(M5:N85)", 5, 86)).toBe("=SUM(M5:N85)");
    expect(pplSyncDongTong("=ROUND(A1,2)", 5, 86)).toBe("=ROUND(A1,2)");
    expect(pplSyncDongTong("", 5, 86)).toBe("");
  });
});

describe("pplSyncTimDoanMau", () => {
  const MARKS = ["#C6EFCE", "#FFEB9C"];
  const makeReader = (colors) => {
    const calls = [];
    const read = (start, count) => {
      calls.push([start, count]);
      const slice = colors.slice(start, start + count);
      return slice.every((c) => c === slice[0]) ? slice[0] : null;
    };
    return { read, calls };
  };

  it("uses a single read when the whole column has one plain color", () => {
    const { pplSyncTimDoanMau } = getTestables();
    const { read, calls } = makeReader(Array(8).fill("#FFFFFF"));

    expect(pplSyncTimDoanMau(read, 0, 8, MARKS)).toEqual([]);
    expect(calls).toHaveLength(1);
  });

  it("finds separate marked segments case-insensitively", () => {
    const { pplSyncTimDoanMau } = getTestables();
    const colors = ["#FFFFFF", "#c6efce", "#C6EFCE", "#FFFFFF", "#FFFFFF", "#FFEB9C", "#FFFFFF", "#FFFFFF"];
    const { read } = makeReader(colors);

    expect(pplSyncTimDoanMau(read, 0, 8, MARKS)).toEqual([
      { start: 1, count: 2 },
      { start: 5, count: 1 },
    ]);
  });

  it("treats empty color as mixed and keeps user colors", () => {
    const { pplSyncTimDoanMau } = getTestables();
    const colors = ["#800080", "#C6EFCE"];
    const read = (start, count) => {
      const slice = colors.slice(start, start + count);
      return slice.every((c) => c === slice[0]) ? slice[0] : "";
    };

    expect(pplSyncTimDoanMau(read, 0, 2, MARKS)).toEqual([{ start: 1, count: 1 }]);
  });

  it("returns nothing for an empty region", () => {
    const { pplSyncTimDoanMau } = getTestables();

    expect(pplSyncTimDoanMau(() => "#FFFFFF", 0, 0, MARKS)).toEqual([]);
  });
});

describe("pplSyncTomTat", () => {
  const meta = { listNo: "000631", status: "active" };

  it("reports added rows with color legend", () => {
    const { pplSyncLapKeHoach, pplSyncTomTat } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6007-1A", 911111)]), [
      systemLine(1, "EUF6007-1A", 911111),
      systemLine(2, "EUF6007-1A", 984945, { product_name: "Tủ gỗ 2 cánh" }),
    ]);

    const text = pplSyncTomTat(plan, meta, 4, 50);

    expect(text).toContain("Đồng bộ phiếu 000631 (active) xong trong 4s.");
    expect(text).toContain("Thêm 1 · Cập nhật 0 · Xoá 0");
    expect(text).toContain("+ 984945 · EUF6007-1A · Tủ gỗ 2 cánh");
    expect(text).toContain("Chú thích: nền xanh lá nhạt = dòng bổ sung, nền vàng nhạt = dòng thay đổi.");
  });

  it("reports no change for an empty plan", () => {
    const { pplSyncLapKeHoach, pplSyncTomTat } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6007-1A", 911111)]), [
      systemLine(1, "EUF6007-1A", 911111),
    ]);

    expect(pplSyncTomTat(plan, meta, 1, 50)).toContain("Đã khớp với hệ thống, không có thay đổi.");
  });

  it("warns about non-editable list status", () => {
    const { pplSyncLapKeHoach, pplSyncTomTat } = getTestables();
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6007-1A", 911111)]), [
      systemLine(1, "EUF6007-1A", 911111),
    ]);

    expect(pplSyncTomTat(plan, { listNo: "000631", status: "converted" }, 1, 50)).toContain(
      "Lưu ý: phiếu đang ở trạng thái converted, dữ liệu chỉ để đối chiếu.",
    );
  });

  it("truncates long lists", () => {
    const { pplSyncLapKeHoach, pplSyncTomTat } = getTestables();
    const lines = [systemLine(1, "EUF6007-1A", 911111)];
    for (let i = 0; i < 5; i++) lines.push(systemLine(2 + i, "EUF6007-1A", 950000 + i));
    const plan = pplSyncLapKeHoach(readSheet([dataRow(1, "EUF6007-1A", 911111)]), lines);

    const text = pplSyncTomTat(plan, meta, 1, 2);

    expect(text).toContain("… và 3 dòng khác");
  });
});
