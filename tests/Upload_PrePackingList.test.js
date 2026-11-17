import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__UPLOAD_PRE_PACKING_TESTABLES__;

  if (!testables) {
    throw new Error("Upload_PrePackingList test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../finished-goods/Upload_PrePackingList.ts");
});

describe("Upload_PrePackingList meta parsing", () => {
  it("extracts list number from NO cell text", () => {
    const { prePackTrichXuatSoPhieu } = getTestables();

    expect(prePackTrichXuatSoPhieu("NO:  000463")).toBe("000463");
    expect(prePackTrichXuatSoPhieu("NO：000464")).toBe("000464");
    expect(prePackTrichXuatSoPhieu("000465")).toBe("000465");
  });

  it("parses Vietnamese and Chinese date formats", () => {
    const { prePackChuanHoaNgayListDate, prePackParseMotDongNgay } = getTestables();

    expect(prePackParseMotDongNgay("27 tháng 08 năm 2026")).toBe("2026-08-27");
    expect(prePackParseMotDongNgay("2026年08月27")).toBe("2026-08-27");
    expect(prePackChuanHoaNgayListDate("27 tháng 08 năm 2026\n2026年08月27")).toBe(
      "2026-08-27",
    );
    expect(prePackChuanHoaNgayListDate("27/08/2026")).toBe("2026-08-27");
    expect(prePackChuanHoaNgayListDate("2026-08-27")).toBe("2026-08-27");
  });

  it("parses shortage API messages into shortage rows", () => {
    const { prePackParseShortageMessage, prePackBuildShortageRows } = getTestables();

    expect(
      prePackParseShortageMessage(
        "Tồn thực tế SKU 911111 / HĐ EUM2026-NEG chỉ còn 5, không đủ để trừ 20.",
      ),
    ).toEqual({
      sku: "911111",
      contract_no: "EUM2026-NEG",
      product_name: "",
      available: 5,
      required: 20,
    });

    const rows = prePackBuildShortageRows(
      [
        "Tồn thực tế SKU 911111 / HĐ EUM2026-NEG chỉ còn 5, không đủ để trừ 20.",
        "Tồn thực tế SKU 922222 / HĐ EUM2026-NEG chỉ còn 5, không đủ để trừ 15.",
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
      ],
    );

    expect(rows).toHaveLength(2);
    expect(rows[0].product_name).toBe("SKU A");
    expect(rows[1].sku).toBe("922222");
  });
});

describe("Upload_PrePackingList system line guard", () => {
  it("formats the 409 system_lines_missing response with guidance", () => {
    const { prePackFormatMissingSystemItems } = getTestables();

    const text = prePackFormatMissingSystemItems({
      message: "Danh sách 000631 có 2 SKU bổ sung trên hệ thống không có trong file Excel.",
      code: "system_lines_missing",
      missing_items: [
        { sku: "984945", contract_no: "EUF6007-1A", product_name: "Tủ gỗ 2 cánh" },
        { sku: "912345A", contract_no: "EUF6001-2B", product_name: "" },
      ],
    });

    expect(text).toContain(
      "LỖI HTTP 409: Danh sách 000631 có 2 SKU bổ sung trên hệ thống không có trong file Excel.",
    );
    expect(text).toContain("  - 984945 · EUF6007-1A · Tủ gỗ 2 cánh");
    expect(text).toContain("  - 912345A · EUF6001-2B");
    expect(text).toContain("Sync_PrePackingList");
    expect(text).toContain("replaceAll: true");
  });

  it("lists removed system items after a replace_all upload", () => {
    const { prePackFormatRemovedSystemItems } = getTestables();

    expect(prePackFormatRemovedSystemItems({ removed_system_items: [] })).toEqual([]);
    expect(prePackFormatRemovedSystemItems({})).toEqual([]);
    expect(
      prePackFormatRemovedSystemItems({
        removed_system_items: [
          { sku: "984945", contract_no: "EUF6007-1A", product_name: "Tủ gỗ 2 cánh" },
        ],
      }),
    ).toEqual([
      "Đã loại bỏ 1 SKU bổ sung trên hệ thống:",
      "  - 984945 · EUF6007-1A · Tủ gỗ 2 cánh",
    ]);
  });
});
