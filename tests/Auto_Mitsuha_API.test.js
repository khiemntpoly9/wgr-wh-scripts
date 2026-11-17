import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__AUTO_MITSUHA_API_TESTABLES__;
  if (!testables) throw new Error("Auto_Mitsuha_Super test helpers were not loaded.");
  return testables;
};

beforeAll(async () => {
  await import("../Auto_Mitsuha_Super.ts");
});

describe("Auto_Mitsuha_Super helpers", () => {
  it("adds the dropped dispatch invoice after column S and the order number after it", () => {
    const { mitsuhaApiTaoDong } = getTestables();
    const row = mitsuhaApiTaoDong({
      contract_no: "EUF2026-001",
      sku: "912345",
      pair_key: "912345|EUF2026-001",
      base_sku: "912345",
      product_name: "Test",
      dimensions: "40*30*20",
      order_quantity: 100,
      unit: "thùng",
      previously_shipped_quantity: 20,
      requested_quantity: 40,
      pack_unit: "thùng",
      batch_no: "LO-01",
      unit_weight: 2.5,
      total_weight: 100,
      unit_volume: 0.024,
      total_volume: 0.96,
      location_notes: "A1",
      location_code: "A1",
      classification_code: "F",
      current_stock: 45,
      pending_production: -10,
      dropped_quantity: 8,
      dropped_label: "RỚT 8/40",
      dropped_dispatch_reference: "PXK-001 — 03/09/2026",
      dropped_invoice_no: "INV-001",
      order_no: "WGF25003EU WGF25004EU",
      latest_receipt_date: "2026-09-05",
      eligible: false,
      warning_reason: "Chưa đủ cặp",
    }, 1);

    expect(row).toHaveLength(26);
    expect(row.slice(0, 9)).toEqual([1, "EUF2026-001", "912345", "Test", "40*30*20", 100, "PCS", 20, 40]);
    expect(row[9]).toBe("箱/ THÙNG CARTON");
    expect(row[15]).toBe("A1");
    expect(row[17]).toBe("RỚT 8/40");
    expect(row[18]).toBe("PXK-001 — 03/09/2026");
    expect(row[19]).toBe("INV-001");
    expect(row[20]).toBe("WGF25003EU WGF25004EU");
    expect(row[21]).toBe("F");
    expect(row[22]).toBe(45);
    expect(row[23]).toBe(-10);
    expect(row[24]).toBe("Chưa đủ cặp");
    expect(row[25]).toBe(46270);
  });

  it("chọn PACKAGE khi chiều cao thứ ba bằng 1,1", () => {
    const { mitsuhaApiChieuCaoKichThuoc, mitsuhaApiDonViDongGoi, mitsuhaApiTaoDong } =
      getTestables();
    expect(mitsuhaApiChieuCaoKichThuoc("26,5*37*1,1")).toBe(1.1);
    expect(mitsuhaApiDonViDongGoi("26,5*37*1,1")).toBe("PACKAGE");
    expect(mitsuhaApiDonViDongGoi("37,5*26,5*1,1")).toBe("PACKAGE");
    expect(mitsuhaApiDonViDongGoi("37.5 x 26.5 x 1.1")).toBe("PACKAGE");
    expect(mitsuhaApiDonViDongGoi("40*30*20")).toBe("箱/ THÙNG CARTON");

    const row = mitsuhaApiTaoDong(
      {
        contract_no: "EUF2026-001",
        sku: "912345",
        pair_key: "912345|EUF2026-001",
        base_sku: "912345",
        product_name: "Test",
        dimensions: "26,5*37*1,1",
        order_quantity: 100,
        unit: "thùng",
        previously_shipped_quantity: 20,
        requested_quantity: 40,
        pack_unit: "thùng",
        batch_no: "LO-01",
        unit_weight: 2.5,
        total_weight: 100,
        unit_volume: 0.024,
        total_volume: 0.96,
        location_notes: "A1",
        location_code: "A1",
        classification_code: "F",
        current_stock: 45,
        pending_production: -10,
        eligible: true,
      },
      1,
    );
    expect(row[9]).toBe("PACKAGE");
  });

  it("creates a valid sheet name and Excel date serial", () => {
    const { mitsuhaApiNgayExcel, mitsuhaApiTenSheet, mitsuhaApiKhoaSkuHopDong } = getTestables();
    expect(mitsuhaApiTenSheet("")).toBe("pre-picking-list");
    expect(mitsuhaApiTenSheet(undefined)).toBe("pre-picking-list");
    expect(mitsuhaApiTenSheet("  soan-hang-tuan-42 ")).toBe("soan-hang-tuan-42");
    expect(mitsuhaApiKhoaSkuHopDong(" 912345 ", " EUF 001 ")).toBe("912345|EUF001");
    expect(mitsuhaApiNgayExcel("2026-09-05")).toBe(46270);
  });

  it("merges backend Excel output settings over defaults", () => {
    const { mitsuhaApiCauHinhExcel } = getTestables();
    expect(mitsuhaApiCauHinhExcel(undefined).start_row).toBe(2);

    const output = mitsuhaApiCauHinhExcel({
      sheet_name: "PPL UE",
      start_row: 3,
      clear_old_data: false,
      warning_fill_color: "#FF0000",
      duplicate_sku_fill_color: "yellow",
    });
    expect(output.sheet_name).toBe("PPL UE");
    expect(output.start_row).toBe(3);
    expect(output.clear_old_data).toBe(false);
    expect(output.enable_excel_filter).toBe(true);
    expect(output.warning_fill_color).toBe("#FF0000");
    expect(output.duplicate_sku_fill_color).toBe("#D9E1F2");
    expect(mitsuhaApiCauHinhExcel({ start_row: 1 }).start_row).toBe(2);
  });

  it("rejects URLs unavailable to Office Scripts", () => {
    const { mitsuhaApiKiemTraUrl } = getTestables();
    expect(mitsuhaApiKiemTraUrl("http://localhost:8000/api/test")).toContain("HTTPS");
    expect(mitsuhaApiKiemTraUrl("https://app.wgr-wh.site/api/v1/test")).toBeNull();
  });
});
