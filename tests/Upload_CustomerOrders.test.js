import { beforeAll, describe, expect, it } from "vitest";
import { createMockExcelScriptEnums } from "./helpers/excelMocks.js";

const getTestables = () => {
  const testables = globalThis.__UPLOAD_CUSTOMER_ORDERS_TESTABLES__;
  if (!testables) throw new Error("Upload_CustomerOrders helpers were not loaded.");
  return testables;
};

beforeAll(async () => {
  globalThis.ExcelScript = createMockExcelScriptEnums();
  await import("../Upload_CustomerOrders.ts");
});

describe("Upload_CustomerOrders", () => {
  it("reads from row 2 and groups lines by order number", () => {
    const { customerOrderUpReadRows } = getTestables();
    const result = customerOrderUpReadRows([
      ["WGF25003E1", "901251", "UAS5001-1", "Tủ A", 6],
      ["WGF25003E1", "901251", "EUF5002-1", "Tủ A", 4],
      ["WGF25004B", "901252", "JPT4001-2", "Tủ B", 1200],
    ]);

    expect(result.errors).toEqual([]);
    expect(result.orders).toHaveLength(2);
    expect(result.orders[0].order_no).toBe("WGF25003E1");
    expect(result.orders[0].items).toHaveLength(2);
    expect(result.orders[1].items[0].ordered_quantity).toBe(1200);
  });

  it("keeps duplicate SKU + contract rows separate and reports their Excel rows", () => {
    const {
      customerOrderUpReadRows,
      customerOrderUpDuplicateNotes,
      customerOrderUpStripLocalFields,
    } = getTestables();
    const result = customerOrderUpReadRows([
      ["WGF26001E1", "990088", "EUF5006-1", "Tủ A", 30],
      ["WGF26001E1", "990088", "EUF5006-2", "Tủ A", 15],
      ["WGF26001E1", "990088", "EUF5006-1", "Tủ A", 40],
      ["WGF26002E1", "990088", "EUF5006-1", "Tủ A", 25],
    ]);

    expect(result.orders[0].items.map((item) => item.ordered_quantity)).toEqual([
      30, 15, 40,
    ]);

    const notes = customerOrderUpDuplicateNotes(result.orders);
    expect(notes).toHaveLength(1);
    expect(notes[0]).toContain("WGF26001E1");
    expect(notes[0]).toContain("dòng 2: 30");
    expect(notes[0]).toContain("dòng 4: 40");

    const [stripped] = customerOrderUpStripLocalFields(result.orders);
    expect(stripped.items[0]).toEqual({
      sku: "990088",
      contract_no: "EUF5006-1",
      product_name: "Tủ A",
      ordered_quantity: 30,
    });
  });

  it("reports the exact Excel row for incomplete or invalid data", () => {
    const { customerOrderUpReadRows } = getTestables();
    const result = customerOrderUpReadRows([
      ["WGF25003E1", "", "UAS5001-1", "Tủ A", 6],
      ["WGF25003E1", "901251", "UAS5001-1", "Tủ A", 1.5],
    ]);

    expect(result.orders).toEqual([]);
    expect(result.errors[0]).toContain("Dòng 2");
    expect(result.errors[0]).toContain("cột B");
    expect(result.errors[1]).toContain("Dòng 3");
    expect(result.errors[1]).toContain("số nguyên dương");
  });

  it("ignores fully blank rows", () => {
    const { customerOrderUpReadRows } = getTestables();
    const result = customerOrderUpReadRows([
      ["", "", "", "", ""],
      ["WGF25003E1", "901251", "UAS5001-1", "Tủ A", 6],
    ]);

    expect(result.dataRowCount).toBe(1);
    expect(result.orders).toHaveLength(1);
  });

  it("accepts and groups rows without an order number", () => {
    const { customerOrderUpReadRows } = getTestables();
    const result = customerOrderUpReadRows([
      ["", "901251", "UAS5001-1", "Tủ A", 6],
      ["", "901252", "UAS5001-1", "Tủ B", 4],
    ]);

    expect(result.errors).toEqual([]);
    expect(result.orders).toHaveLength(1);
    expect(result.orders[0].order_no).toBeNull();
    expect(result.orders[0].items).toHaveLength(2);
  });

  it("merges chunk reads and builds bulk batches under item limits", () => {
    const {
      customerOrderUpMergeOrders,
      customerOrderUpSplitForUpload,
      customerOrderUpBuildBatches,
      customerOrderUpValidateUrl,
    } = getTestables();

    const orderMap = new Map();
    customerOrderUpMergeOrders(orderMap, [
      {
        order_no: "WGF25003E1",
        status: "confirmed",
        items: [
          {
            sku: "901251",
            contract_no: "UAS5001-1",
            product_name: "Tủ A",
            ordered_quantity: 6,
          },
        ],
      },
    ]);
    customerOrderUpMergeOrders(orderMap, [
      {
        order_no: "WGF25003E1",
        status: "confirmed",
        items: [
          {
            sku: "901251",
            contract_no: "EUF5002-1",
            product_name: "Tủ A",
            ordered_quantity: 4,
          },
        ],
      },
      {
        order_no: null,
        status: "confirmed",
        items: Array.from({ length: 5 }, (_, i) => ({
          sku: `90126${i}`,
          contract_no: `UAS501${i}-1`,
          product_name: `SP ${i}`,
          ordered_quantity: 1,
        })),
      },
    ]);

    expect(orderMap.get("WGF25003E1").items).toHaveLength(2);

    const { uploadOrders, splitNotes } = customerOrderUpSplitForUpload(
      Array.from(orderMap.values()),
      2,
    );
    expect(splitNotes).toEqual([]);
    expect(uploadOrders.filter((order) => order.order_no === null)).toHaveLength(
      3,
    );

    const batches = customerOrderUpBuildBatches(uploadOrders, 2, 10);
    expect(batches.length).toBeGreaterThan(1);
    expect(
      customerOrderUpValidateUrl(
        "https://app.example.com/api/warehouse/customer-orders/bulk-import",
      ),
    ).toBeNull();
  });

  it("splits oversized named orders into create + append chunks", () => {
    const { customerOrderUpSplitForUpload } = getTestables();
    const items = Array.from({ length: 5 }, (_, i) => ({
      sku: `90130${i}`,
      contract_no: `UAS60${i}-1`,
      product_name: `SP ${i}`,
      ordered_quantity: 1,
    }));

    const { uploadOrders, splitNotes } = customerOrderUpSplitForUpload(
      [{ order_no: "WGF25999E1", status: "confirmed", items }],
      2,
    );

    expect(splitNotes).toHaveLength(1);
    expect(uploadOrders).toHaveLength(3);
    expect(uploadOrders[0].append_to_existing).toBe(false);
    expect(uploadOrders[0].items).toHaveLength(2);
    expect(uploadOrders[1].append_to_existing).toBe(true);
    expect(uploadOrders[1].order_no).toBe("WGF25999E1");
    expect(uploadOrders[2].append_to_existing).toBe(true);
    expect(uploadOrders[2].items).toHaveLength(1);
  });

  it("accepts a sheet with matching headers or any data from row 2", () => {
    const { customerOrderUpSheetHopLe, customerOrderUpNormalizeHeader } =
      getTestables();

    expect(customerOrderUpNormalizeHeader("Mã Đơn Hàng")).toContain("ma don");
    expect(
      customerOrderUpSheetHopLe(
        ["Mã đơn hàng", "SKU", "Mã hợp đồng", "Tên hàng", "Số lượng đơn hàng"],
        [],
      ),
    ).toBe(true);
    expect(
      customerOrderUpSheetHopLe(["A", "B", "C", "D", "E"], []),
    ).toBe(false);
    expect(
      customerOrderUpSheetHopLe(undefined, [
        ["WGF1", "901251", "UAS1", "Tủ", 1],
      ]),
    ).toBe(true);
    expect(customerOrderUpSheetHopLe(undefined, [["", "", "", "", ""]])).toBe(
      false,
    );
  });

  it("writes sample header and demo rows when creating the template sheet", () => {
    const { customerOrderUpTaoSheetMau } = getTestables();
    const writes = [];
    const formats = [];
    const sheet = {
      setVisibility() {},
      getUsedRange() {
        return { clear() {} };
      },
      getRange(address) {
        return {
          setValues(values) {
            writes.push({ address, values });
          },
          getFormat() {
            return {
              getFont() {
                return { setBold(v) { formats.push(["bold", v]); } };
              },
              setHorizontalAlignment(v) {
                formats.push(["align", v]);
              },
              getFill() {
                return { setColor(v) { formats.push(["fill", v]); } };
              },
              autofitColumns() {
                formats.push(["autofit"]);
              },
            };
          },
        };
      },
      getFreezePanes() {
        return { freezeRows(n) { formats.push(["freeze", n]); } };
      },
      activate() {
        formats.push(["activate"]);
      },
    };
    const workbook = {
      getWorksheet(name) {
        return name === "Đơn đặt hàng" ? sheet : undefined;
      },
      addWorksheet() {
        return sheet;
      },
    };

    const name = customerOrderUpTaoSheetMau(
      workbook,
      "Đơn đặt hàng",
      ["Mã đơn hàng", "SKU", "Mã hợp đồng", "Tên hàng", "Số lượng đơn hàng"],
      [["WGF1", "901251", "UAS1", "Tủ A", 6]],
    );

    expect(name).toBe("Đơn đặt hàng");
    expect(writes[0].address).toBe("A1:E2");
    expect(writes[0].values[0][1]).toBe("SKU");
    expect(writes[0].values[1][0]).toBe("WGF1");
    expect(formats).toContainEqual(["activate"]);
  });
});
