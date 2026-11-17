import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__UPLOAD_PXK_TESTABLES__;

  if (!testables) {
    throw new Error("Upload_PXK test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Upload_PXK.ts");
});

describe("Upload_PXK meta parsing", () => {
  it("extracts dispatch number from NO cell text", () => {
    const { pxkUpTrichXuatSoPhieu } = getTestables();

    expect(pxkUpTrichXuatSoPhieu("NO:  000462")).toBe("000462");
    expect(pxkUpTrichXuatSoPhieu("NO：000464")).toBe("000464");
    expect(pxkUpTrichXuatSoPhieu("000465")).toBe("000465");
  });

  it("extracts container and invoice codes", () => {
    const { pxkUpTrichXuatMa } = getTestables();

    expect(pxkUpTrichXuatMa("CSNU7515625", ["CONT", "CONTAINER"])).toBe(
      "CSNU7515625",
    );
    expect(
      pxkUpTrichXuatMa("CONT: CSNU7515625", ["CONT", "CONTAINER"]),
    ).toBe("CSNU7515625");
    expect(pxkUpTrichXuatMa("WGR-20260095", ["INVOICE", "INV"])).toBe(
      "WGR-20260095",
    );
    expect(pxkUpTrichXuatMa("INVOICE: WGR-20260095", ["INVOICE", "INV"])).toBe(
      "WGR-20260095",
    );
  });

  it("parses Vietnamese and Chinese date formats from D3", () => {
    const { pxkUpChuanHoaNgay, pxkUpParseMotDongNgay } = getTestables();

    expect(pxkUpParseMotDongNgay("04 tháng 08 năm 2026")).toBe("2026-08-04");
    expect(pxkUpParseMotDongNgay("2026年09月04")).toBe("2026-09-04");
    expect(pxkUpParseMotDongNgay("2026年09月04,")).toBe("2026-09-04");
    expect(pxkUpParseMotDongNgay("2026年09月04日")).toBe("2026-09-04");
    expect(
      pxkUpChuanHoaNgay("04 tháng 08 năm 2026\n2026年09月04,"),
    ).toBe("2026-08-04");
    expect(pxkUpChuanHoaNgay("04/08/2026")).toBe("2026-08-04");
    expect(pxkUpChuanHoaNgay("2026-08-04")).toBe("2026-08-04");
  });

  it("reads Auto_PXK sale rows into upload payload items", () => {
    const { pxkUpDocDuLieu } = getTestables();

    const rows = [
      ["header"],
      ["STT\n序号", "HĐ", "SKU", "Tên", "KT", "SL", "ĐVT", "Đã", "Đóng", "ĐVT2", "Lô", "TL", "TTL", "TT", "TTT", "GC", "", "Y"],
      [1, "EUM2026-001", "912345ABC", "Bàn", "100*50*30", 10, "PCS", 0, 8, "箱", "2026080401", 12.5, 100, 0.1, 0.8, "note", "", "Y"],
      ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "SUM", "", "=SUM(F3:F3)", "", "", "", "", "", "", "", "", "", "", "", ""],
    ];

    const items = pxkUpDocDuLieu(rows, ["STT", "序号"]);
    expect(items).toHaveLength(1);
    expect(items[0].sku).toBe("912345");
    expect(items[0].contract_no).toBe("EUM2026-001");
    expect(items[0].dispatch_quantity).toBe(10);
    expect(items[0].packed_quantity).toBe(8);
    expect(items[0].batch_no).toBe("2026080401");
    expect(items[0].pmc_approved).toBe("Y");
  });

  it("keeps every leading digit of 7-digit SKUs", () => {
    const { pxkUpDocDuLieu } = getTestables();
    const line = (stt, sku) => [stt, "EUM2026-001", sku, "Bàn", "", 5, "PCS", 0, 5];
    const rows = [
      ["STT", "HĐ", "SKU", "Tên"],
      line(1, 9123456),
      line(2, "9123457"),
      line(3, "9123458-01"),
      line(4, "9123459ABC"),
      line(5, "9123450a"),
      line(6, " 9123451 "),
    ];

    const items = pxkUpDocDuLieu(rows, ["STT"]);
    expect(items.map((item) => item.sku)).toEqual([
      "9123456",
      "9123457",
      "9123458",
      "9123459",
      "9123450A",
      "9123451",
    ]);
  });

  it("parses shortage messages and enriches product names from upload rows", () => {
    const { pxkUpParseShortageMessage, pxkUpBuildShortageRows } = getTestables();

    expect(
      pxkUpParseShortageMessage(
        "Tồn thực tế SKU 911111 / HĐ EUM2026-NEG chỉ còn 5, không đủ để trừ 20.",
      ),
    ).toEqual({
      sku: "911111",
      contract_no: "EUM2026-NEG",
      product_name: "",
      available: 5,
      required: 20,
    });

    const rows = pxkUpBuildShortageRows(
      [
        "Tồn thực tế SKU 911111 / HĐ EUM2026-NEG chỉ còn 5, không đủ để trừ 20.",
        "Tồn thực tế SKU 922222 / HĐ EUM2026-NEG chỉ còn 0, không đủ để trừ 15.",
      ],
      [
        {
          stt: 1,
          contract_no: "EUM2026-NEG",
          sku: "911111",
          product_name: "SKU A",
          dimensions: "",
          dispatch_quantity: 20,
          unit: "",
          previously_shipped_quantity: 0,
          packed_quantity: 20,
          pack_unit: "",
          batch_no: "",
          unit_weight: null,
          total_weight: null,
          unit_volume: null,
          total_volume: null,
          location_notes: "",
          notes: "",
          pmc_approved: "",
        },
        {
          stt: 2,
          contract_no: "EUM2026-NEG",
          sku: "922222",
          product_name: "SKU B",
          dimensions: "",
          dispatch_quantity: 15,
          unit: "",
          previously_shipped_quantity: 0,
          packed_quantity: 15,
          pack_unit: "",
          batch_no: "",
          unit_weight: null,
          total_weight: null,
          unit_volume: null,
          total_volume: null,
          location_notes: "",
          notes: "",
          pmc_approved: "",
        },
      ],
    );

    expect(rows).toEqual([
      {
        sku: "911111",
        contract_no: "EUM2026-NEG",
        product_name: "SKU A",
        available: 5,
        required: 20,
      },
      {
        sku: "922222",
        contract_no: "EUM2026-NEG",
        product_name: "SKU B",
        available: 0,
        required: 15,
      },
    ]);
  });
});
