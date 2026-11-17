import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__UPLOAD_MITSUHA_PRE_PACKING_TESTABLES__;

  if (!testables) {
    throw new Error("Upload_PrePackingList_Mitsuha test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Upload_PrePackingList_Mitsuha.ts");
});

const buildHeader = () => {
  const header = new Array(18).fill("");
  header[3] = "Mã hợp đồng";
  header[4] = "SKU";
  header[7] = "Số lượng đơn hàng";
  header[9] = "Đã xuất trước";
  header[10] = "Số lượng soạn";
  header[17] = "Vị trí";
  return header;
};

const buildRow = (ngay, phieu, hd, sku, slDon, soan, viTri = "A15 B25") => {
  const row = new Array(18).fill("");
  row[0] = ngay;
  row[1] = phieu;
  row[2] = 1;
  row[3] = hd;
  row[4] = sku;
  row[5] = "Bàn làm việc";
  row[6] = "100*50*30";
  row[7] = slDon;
  row[9] = 0;
  row[10] = soan;
  row[17] = viTri;
  return row;
};

describe("Upload_PrePackingList_Mitsuha parsing", () => {
  it("normalizes list number and date key from columns A/B", () => {
    const { mitsuhaChuanHoaSoPhieu, mitsuhaNgayKey } = getTestables();

    expect(mitsuhaChuanHoaSoPhieu(434)).toBe("434");
    expect(mitsuhaChuanHoaSoPhieu("NO: 000463")).toBe("000463");
    expect(mitsuhaNgayKey("14/08/2026")).toBe("2026-08-14");
  });

  it("normalizes warehouse location in column R", () => {
    const { mitsuhaChuanHoaViTri } = getTestables();

    expect(mitsuhaChuanHoaViTri(" b044-c72 ")).toBe("B44 C72");
    expect(mitsuhaChuanHoaViTri("B25 A15")).toBe("A15 B25");
  });

  it("reads rows using Mitsuha A→R layout and keeps only latest date", () => {
    const { mitsuhaDocDuLieu } = getTestables();

    const header = buildHeader();
    const rows = [
      header,
      buildRow("12/08/2026", 429, "EU-01", "900872", 15, 15),
      buildRow("14/08/2026", 434, "EU-01", "901019A", 8, 8, "A15"),
      buildRow("14/08/2026", 435, "EU-01", "901019A", 8, 4, "B25"),
    ];

    const batches = mitsuhaDocDuLieu(rows, "phieu_moi_nhat");

    expect(batches).toHaveLength(1);
    expect(batches[0].listDateIso).toBe("2026-08-14");
    expect(batches[0].listNo).toBe("435");
    expect(batches[0].items).toHaveLength(2);
    expect(batches[0].items[0].contract_no).toBe("EU-01");
    expect(batches[0].items[0].sku).toBe("901019");
    expect(batches[0].items[0].dispatch_quantity).toBe(8);
    expect(batches[0].items[0].packed_quantity).toBe(8);
    expect(batches[0].items[0].location_notes).toBe("A15");
    expect(batches[0].items[1].packed_quantity).toBe(4);
    expect(batches[0].items[1].location_notes).toBe("B25");
  });

  it("groups all slips by date and list number in tat_ca mode", () => {
    const { mitsuhaDocDuLieu } = getTestables();

    const header = buildHeader();
    const rows = [
      header,
      buildRow("12/08/2026", 429, "EU-01", "900872", 15, 15),
      buildRow("14/08/2026", 434, "EU-01", "901019A", 8, 8, "A15"),
      buildRow("14/08/2026", 435, "EU-01", "901020B", 10, 6, "B25"),
    ];

    const batches = mitsuhaDocDuLieu(rows, "tat_ca");

    expect(batches).toHaveLength(3);
    expect(batches[0].listNo).toBe("429");
    expect(batches[0].listDateIso).toBe("2026-08-12");
    expect(batches[0].items).toHaveLength(1);
    expect(batches[1].listNo).toBe("434");
    expect(batches[1].listDateIso).toBe("2026-08-14");
    expect(batches[2].listNo).toBe("435");
    expect(batches[2].items[0].sku).toBe("901020");
  });

  it("finds header row when column E is SKU and column D is contract", () => {
    const { mitsuhaTimDongTieuDe } = getTestables();

    const rows = [
      ["", "", "", "", ""],
      buildHeader(),
      buildRow("14/08/2026", 434, "EU-01", "901019A", 8, 8),
    ];

    expect(mitsuhaTimDongTieuDe(rows)).toBe(1);
  });
});
