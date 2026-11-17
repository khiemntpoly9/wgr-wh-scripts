import { beforeAll, describe, expect, it } from "vitest";

import { createMockExcelScriptEnums } from "./helpers/excelMocks.js";

const getTestables = () => {
  const testables = globalThis.__AUTO_PXK_SUPER_TESTABLES__;
  if (!testables) throw new Error("Auto_PXK_Super test helpers were not loaded.");
  return testables;
};

beforeAll(async () => {
  globalThis.ExcelScript = createMockExcelScriptEnums();
  await import("../finished-goods/Auto_PXK_Super.ts");
});

describe("Auto_PXK_Super helpers", () => {
  it("chọn PACKAGE khi chiều cao thứ ba bằng 1,1", () => {
    const { pxkSuperDimensionHeight, pxkSuperPackUnit } = getTestables();
    expect(pxkSuperDimensionHeight("44*24*1,1")).toBe(1.1);
    expect(pxkSuperPackUnit("44 x 24 x 1.1")).toBe("PACKAGE");
    expect(pxkSuperPackUnit("42*24.5*25")).toBe("箱/ THÙNG CARTON");
  });

  it("tạo dòng PXK A-S với PCS và đơn vị đóng gói tự động", () => {
    const { pxkSuperLoadedRow, pxkSuperDroppedRow } = getTestables();
    const loaded = pxkSuperLoadedRow({
      contract_no: "EU-01",
      sku: "912345",
      dimensions: "44*24*1,1",
      dispatch_quantity: 20,
      packed_quantity: 20,
      unit_weight: 2.5,
      unit_volume: 0.01,
      location_code: "A10 B3",
      notes: "Hàng bù",
    });
    expect(loaded).toHaveLength(19);
    expect(loaded[6]).toBe("PCS");
    expect(loaded[9]).toBe("PACKAGE");
    expect(loaded[12]).toBe("");
    expect(loaded[13]).toBe(""); // N để trống — điền công thức từ E
    expect(loaded[14]).toBe("");
    expect(loaded[15]).toBe("A10 B3"); // P: vị trí
    expect(loaded[16]).toBe("Hàng bù"); // Q: ghi chú
    expect(loaded[17]).toBe(""); // Cột R không điền

    const dropped = pxkSuperDroppedRow(
      {
        requested_quantity: 20,
        loaded_quantity: 12,
        dropped_quantity: 8,
        unit_weight: 2.5,
        unit_volume: 0.01,
        dimensions: "40*30*25",
      },
      1,
    );
    expect(dropped).toEqual([1, "", "", "", 8, 12, 20, 0.08, ""]);
    expect(dropped).toHaveLength(9);

    const droppedWithLocation = pxkSuperDroppedRow(
      { dropped_quantity: 3, location_code: " A7 A10 " },
      2,
    );
    expect(droppedWithLocation[8]).toBe("A7 A10");
  });

  it("xác định các dòng hàng bù và dòng xuất trước đó cần tô nền", () => {
    const {
      pxkSuperCompensationRowIndexes,
      pxkSuperHighlightCompensationRows,
    } = getTestables();
    const colored = [];
    const sheet = {
      getName: () => "sale",
      getRange: (address) => ({
        getFormat: () => ({
          getFill: () => ({
            setColor: (color) => colored.push({ address, color }),
          }),
        }),
      }),
    };
    const items = [
      { sku: "NORMAL" },
      { sku: "COMP", is_compensation: true },
      { sku: "OFFSET", is_compensation_offset: true },
      { sku: "API", highlight_compensation: true },
    ];
    expect(
      pxkSuperCompensationRowIndexes(items),
    ).toEqual([1, 2, 3]);
    pxkSuperHighlightCompensationRows(sheet, 10, items);
    expect(colored).toEqual([
      { address: "A12:S12", color: "#FFF2CC" },
      { address: "A13:S13", color: "#FFF2CC" },
      { address: "A14:S14", color: "#FFF2CC" },
    ]);
  });

  it("bung từng lần nhập SKU thành dòng riêng (không gộp tổng)", () => {
    const { pxkSuperBuildSkuLoadedRows, pxkSuperSkuLoadedRow } = getTestables();
    const rows = pxkSuperBuildSkuLoadedRows([
      {
        contract_no: "EU-01",
        sku: "912345",
        product_name: "Bàn",
        dispatch_quantity: 28,
        packed_quantity: 28,
        entry_quantities: [13, 15],
      },
      {
        contract_no: "EU-02",
        sku: "912999",
        product_name: "Ghế",
        dispatch_quantity: 5,
        entry_quantities: [5],
      },
    ]);
    expect(rows).toEqual([
      {
        contract_no: "EU-01",
        sku: "912345",
        product_name: "Bàn",
        loaded_quantity: 13,
      },
      {
        contract_no: "EU-01",
        sku: "912345",
        product_name: "Bàn",
        loaded_quantity: 15,
      },
      {
        contract_no: "EU-02",
        sku: "912999",
        product_name: "Ghế",
        loaded_quantity: 5,
      },
    ]);
    expect(pxkSuperSkuLoadedRow(rows[0], 1)).toEqual([
      1,
      "EU-01",
      "912345",
      "Bàn",
      13,
    ]);
  });

  it("dùng ticket id làm khóa tracking để chống trùng", () => {
    const { pxkSuperTrackingKey } = getTestables();
    expect(pxkSuperTrackingKey(15, "PXK-01", "CONT", "SEAL")).toBe("ID:15");
    expect(pxkSuperTrackingKey("", "pxk-01", "cont", "seal")).toBe(
      "PXK-01|CONT|SEAL",
    );
  });

  it("bù trọng lượng theo nhóm kích thước và khớp công thức ROUND(L*I,2)", () => {
    const { pxkSuperBalanceWeights } = getTestables();
    const result = pxkSuperBalanceWeights(
      ["40*30*25", "40*30*25"],
      [2, 2],
      [5, 5],
      21,
    );
    expect(result.weights).toEqual([2.1, 2.1]);
    expect(result.total).toBe(21);
    expect(result.difference).toBe(0);
    expect(result.adjusted).toBe(true);
  });

  it("tinh chỉnh cân nặng theo bước 0,01 kg khi có sai số làm tròn", () => {
    const { pxkSuperBalanceWeights } = getTestables();
    const result = pxkSuperBalanceWeights(["A"], [2], [3], 6.03);
    expect(result.weights).toEqual([2.01]);
    expect(result.total).toBe(6.03);
    expect(result.difference).toBe(0);
  });

  it("chỉ chấp nhận API HTTPS public", () => {
    const { pxkSuperValidateUrl } = getTestables();
    expect(pxkSuperValidateUrl("https://app.example.com/api")).toBeNull();
    expect(pxkSuperValidateUrl("http://app.example.com/api")).toContain("HTTPS");
    expect(pxkSuperValidateUrl("https://localhost/api")).toContain("localhost");
  });

  it("ghép mã hợp đồng unique cho chế độ D5 api", () => {
    const { pxkSuperJoinContractsFromItems } = getTestables();
    expect(
      pxkSuperJoinContractsFromItems([
        { contract_no: "EU-01" },
        { contract_no: "eu-01" },
        { contract_no: "EU-02" },
        { contract_no: "" },
      ]),
    ).toBe("EU-01, EU-02");
  });
});
