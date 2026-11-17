import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__UPLOAD_NAMI_TESTABLES__;
  if (!testables) {
    throw new Error("Upload_Nami test helpers were not loaded.");
  }
  return testables;
};

beforeAll(async () => {
  await import("../finished-goods/Upload_Nami.ts");
});

const buildHeader = () => [
  "Nơi nhập",
  "Ngày nhập kho",
  "Mã hợp đồng",
  "Thị trường",
  "Mã phân loại",
  "SKU",
  "Tên hàng",
  "Số lượng đơn hàng",
  "Số lượng đã nhập",
  "Số lượng nhập",
  "Số lượng còn lại",
  "Mã ghép",
  "Trọng lượng",
  "Tổng trọng lượng",
  "Dài",
  "Rộng",
  "Cao",
  "Thể tích",
  "Tổng thể tích",
  "Vị trí",
];

const buildRow = (overrides = {}) => {
  const row = [
    "Nhập kho 1",
    "25/05/2026",
    "JPT4001-2",
    "JP",
    "T",
    "912345",
    "Bàn làm việc",
    100,
    50,
    30,
    20,
    "912345JPT4001-2",
    20,
    600,
    100,
    50,
    30,
    0.165,
    4.95,
    "A05 B44",
  ];
  Object.assign(row, overrides);
  return row;
};

describe("Upload_Nami location parsing", () => {
  it("normalizes warehouse location clusters", () => {
    const { namiUpLayMaViTri } = getTestables();

    expect(namiUpLayMaViTri("A05  B044")).toBe("A5 B44");
    expect(namiUpLayMaViTri("c12")).toBe("C12");
    expect(namiUpLayMaViTri("")).toBe("");
  });

  it("finds the Vị trí column by header name", () => {
    const { namiUpTimChiSoCot } = getTestables();

    expect(namiUpTimChiSoCot(buildHeader(), "vị trí")).toBe(19);
    expect(namiUpTimChiSoCot(["SKU", "Vị trí kho"], "vị trí")).toBe(1);
    expect(namiUpTimChiSoCot(["SKU", "Tên"], "vị trí")).toBe(-1);
  });

  it("prefers displayed text over empty getValues for location", () => {
    const { namiUpLayOChuoi } = getTestables();

    expect(namiUpLayOChuoi(["A5 B44"], [""], 0)).toBe("A5 B44");
    expect(namiUpLayOChuoi([""], ["B3"], 0)).toBe("B3");
    expect(namiUpLayOChuoi(undefined, ["C12"], 0)).toBe("C12");
  });

  it("reads location from column T even when statistics columns follow", () => {
    const { namiUpDocDuLieu } = getTestables();

    const header = [...buildHeader(), "THỐNG KÊ"];
    const row = [...buildRow(), "ignored-stat"];
    const items = namiUpDocDuLieu([[], header, row], 2);

    expect(items).toHaveLength(1);
    expect(items[0].location_code).toBe("A5 B44");
    expect(items[0].sku).toBe("912345");
    expect(items[0].contract_no).toBe("JPT4001-2");
  });

  it("sends SKUs as-is (uppercased, no whitespace) so the backend can match them against orders", () => {
    const { namiUpDocDuLieu } = getTestables();

    const header = buildHeader();
    const skus = [
      "9123456",
      "912345A",
      "912345b",
      " 91234567-01 ",
      "8123456",
    ];
    const rows = skus.map((sku) => buildRow({ 5: sku }));
    const items = namiUpDocDuLieu([[], header, ...rows], 2);

    expect(items.map((item) => item.sku)).toEqual([
      "9123456",
      "912345A",
      "912345B",
      "91234567-01",
      "8123456",
    ]);
  });

  it("reports the Excel row number of each item", () => {
    const { namiUpDocDuLieu } = getTestables();

    const rows = [[], buildHeader(), buildRow(), buildRow({ 5: "912346" })];

    expect(namiUpDocDuLieu(rows, 2).map((item) => item.source_row)).toEqual([
      3, 4,
    ]);
    expect(
      namiUpDocDuLieu(rows, 2, undefined, 4).map((item) => item.source_row),
    ).toEqual([7, 8]);
  });

  it("skips rows whose SL nhập (J) or SL đã nhập (I) is empty, and reports them", () => {
    const { namiUpDocDuLieu } = getTestables();

    const rows = [
      [],
      buildHeader(),
      buildRow(),
      buildRow({ 5: "912346", 9: "" }),
      buildRow({ 5: "912347", 8: "" }),
      buildRow({ 5: "912348", 8: 10, 9: 30 }),
      buildRow({ 5: "912349", 9: 0 }),
    ];
    const skipped = [];
    const items = namiUpDocDuLieu(rows, 2, undefined, 0, skipped);

    expect(items.map((item) => item.sku)).toEqual(["912345"]);
    expect(items[0].total_received).toBe(50);
    expect(skipped.map((r) => [r.source_row, r.sku])).toEqual([
      [4, "912346"],
      [5, "912347"],
      [6, "912348"],
      [7, "912349"],
    ]);
    expect(skipped[1].reason).toContain("cột I");
  });

  it("lists data rows up to the statistics block and spans A..Vị trí for highlighting", () => {
    const { namiUpVungDuLieu } = getTestables();

    const rows = [
      [],
      buildHeader(),
      buildRow(),
      buildRow({ 5: "" }),
      buildRow({ 5: "912346", 9: "" }),
      ["THỐNG KÊ"],
      buildRow({ 5: "912999" }),
    ];

    expect(namiUpVungDuLieu(rows, 2)).toEqual({ dongs: [3, 5], soCot: 20 });
    expect(namiUpVungDuLieu(rows, 2, 4).dongs).toEqual([7, 9]);
    expect(namiUpVungDuLieu([[]], 2)).toEqual({ dongs: [], soCot: 0 });
  });

  it("reads location from getTexts when getValues is empty", () => {
    const { namiUpDocDuLieu } = getTestables();

    const header = buildHeader();
    const valueRow = buildRow({ 19: "" });
    const textRow = buildRow({ 19: "B03" });
    const items = namiUpDocDuLieu([[], header, valueRow], 2, [
      [],
      header,
      textRow,
    ]);

    expect(items).toHaveLength(1);
    expect(items[0].location_code).toBe("B3");
  });
});
