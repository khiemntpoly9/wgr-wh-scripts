import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  const testables = globalThis.__UPLOAD_NAMI_SUPER_TESTABLES__;
  if (!testables) {
    throw new Error("Upload_Nami_Super test helpers were not loaded.");
  }
  return testables;
};

beforeAll(async () => {
  await import("../finished-goods/Upload_Nami_Super.ts");
});

const KY_THANG_5 = { mode: "month", from: "2026-05-01", to: "2026-05-31" };

/** Một dòng "Thành phẩm" A..O; các cột không dùng để trống. */
const buildRow = (overrides = {}) => {
  const row = [
    "", // A
    "", // B
    "912345", // C SKU
    "Bàn làm việc", // D Tên
    "", // E
    100, // F SL đơn hàng
    "", // G
    "JPT4001-2", // H Mã HĐ
    "A05  B44", // I Vị trí
    "25/05/2026", // J Ngày nhập 1
    20, // K SL nhập 1
    "", // L Ngày nhập 2
    "", // M SL nhập 2
    "", // N Ngày nhập 3
    "", // O SL nhập 3
  ];
  Object.assign(row, overrides);
  return row;
};

const asTexts = (rows) => rows.map((row) => row.map((cell) => String(cell)));

describe("Upload_Nami_Super period", () => {
  it("builds day, week and month periods", () => {
    const { namiSuperTinhKy } = getTestables();
    const now = new Date("2026-01-01T00:00:00Z");

    expect(namiSuperTinhKy("25/05/2026", "day", now, 7)).toMatchObject({
      mode: "day",
      from: "2026-05-25",
      to: "2026-05-25",
    });
    expect(namiSuperTinhKy("25/05/2026", "week", now, 7)).toMatchObject({
      from: "2026-05-25",
      to: "2026-05-31",
    });
    expect(namiSuperTinhKy("31/05/2026", "week", now, 7)).toMatchObject({
      from: "2026-05-25",
      to: "2026-05-31",
    });
    expect(namiSuperTinhKy("15/02/2028", "month", now, 7)).toMatchObject({
      mode: "month",
      from: "2028-02-01",
      to: "2028-02-29",
    });
  });

  it("uses today in Vietnam time when no date is configured", () => {
    const { namiSuperTinhKy } = getTestables();

    expect(
      namiSuperTinhKy("", "day", new Date("2026-05-24T18:30:00Z"), 7),
    ).toMatchObject({ from: "2026-05-25", to: "2026-05-25" });
  });

  it("returns null for an invalid date or mode", () => {
    const { namiSuperTinhKy } = getTestables();
    const now = new Date("2026-05-25T00:00:00Z");

    expect(namiSuperTinhKy("32/13/2026", "day", now, 7)).toBeNull();
    expect(namiSuperTinhKy("25/05/2026", "year", now, 7)).toBeNull();
  });
});

describe("Upload_Nami_Super cell parsing", () => {
  it("reads dates like Auto_Nami", () => {
    const { namiSuperLayNgayKey } = getTestables();
    const serial = Date.UTC(2026, 4, 25) / 86400000 + 25569;

    expect(namiSuperLayNgayKey("", serial)).toBe("2026-05-25");
    expect(namiSuperLayNgayKey("25/05/2026", "25/05/2026")).toBe("2026-05-25");
    expect(namiSuperLayNgayKey("", "2026-05-25")).toBe("2026-05-25");
    expect(namiSuperLayNgayKey("05/25/2026", "")).toBe("2026-05-25");
    expect(namiSuperLayNgayKey("25-05-26", "")).toBe("2026-05-25");
    expect(namiSuperLayNgayKey("abc", "abc")).toBeNull();
  });

  it("normalizes warehouse locations", () => {
    const { namiSuperChuanHoaViTri } = getTestables();

    expect(namiSuperChuanHoaViTri("A05  B44")).toBe("A5 B44");
    expect(namiSuperChuanHoaViTri("")).toBe("");
  });
});

describe("Upload_Nami_Super reading receipt rounds", () => {
  it("creates one receipt per round inside the period", () => {
    const { namiSuperDocLanNhap } = getTestables();
    const rows = [
      buildRow({
        2: " 912345 a ",
        11: "26/05/2026",
        12: 30,
        13: "27/05/2026",
        14: 10,
      }),
    ];

    const { lanNhap, boQua, soDongCoNoiDung } = namiSuperDocLanNhap(
      rows,
      asTexts(rows),
      1,
      KY_THANG_5,
    );

    expect(soDongCoNoiDung).toBe(1);
    expect(boQua).toEqual([]);
    expect(lanNhap).toHaveLength(3);
    expect(lanNhap.map((x) => x.round_number)).toEqual([1, 2, 3]);
    expect(lanNhap.map((x) => x.receipt_date)).toEqual([
      "2026-05-25",
      "2026-05-26",
      "2026-05-27",
    ]);
    expect(lanNhap.map((x) => x.quantity)).toEqual([20, 30, 10]);
    expect(lanNhap[0]).toMatchObject({
      source_row: 2,
      sku: "912345A",
      contract_no: "JPT4001-2",
      product_name: "Bàn làm việc",
      order_quantity: 100,
      total_received: 60,
      location_code: "A5 B44",
    });
  });

  it("ignores empty rows and rounds outside the period", () => {
    const { namiSuperDocLanNhap } = getTestables();
    const rows = [
      buildRow({ 2: "", 3: "", 7: "" }),
      buildRow({ 9: "25/04/2026" }),
    ];

    const { lanNhap, boQua, soDongCoNoiDung } = namiSuperDocLanNhap(
      rows,
      asTexts(rows),
      1,
      KY_THANG_5,
    );

    expect(soDongCoNoiDung).toBe(1);
    expect(lanNhap).toEqual([]);
    expect(boQua).toEqual([]);
  });

  it("skips rounds with a bad quantity or missing keys and says why", () => {
    const { namiSuperDocLanNhap } = getTestables();
    const rows = [
      buildRow({ 10: "" }),
      buildRow({ 10: 0 }),
      buildRow({ 10: "abc" }),
      buildRow({ 2: "" }),
      buildRow({ 7: "" }),
      buildRow({ 10: 20, 11: "26/04/2026", 12: -5 }),
    ];

    const { lanNhap, boQua } = namiSuperDocLanNhap(
      rows,
      asTexts(rows),
      4,
      KY_THANG_5,
    );

    expect(lanNhap).toEqual([]);
    expect(boQua.map((x) => [x.source_row, x.round_number])).toEqual([
      [5, 1],
      [6, 1],
      [7, 1],
      [8, 1],
      [9, 1],
      [10, 1],
    ]);
    expect(boQua[0].reason).toContain("SL");
    expect(boQua[3].reason).toContain("SKU");
    expect(boQua[4].reason).toContain("hợp đồng");
    expect(boQua[5].reason).toContain("K + M + O");
  });
});

describe("Upload_Nami_Super spec lookup", () => {
  const rows2026 = [
    ["912345", "Bàn làm việc", "", 20, 100, 50, 30],
    ["", "Ghế gỗ sồi", "", 5, 40, 40, ""],
    ["", "Tủ quần áo gỗ sồi ba cánh", "", 80, 120, 60, 200],
  ];
  const cauHinh = {
    bat: true,
    diemToiThieu: 0.55,
    tyLeTokenKhopToiThieu: 0.5,
    soTokenKhopToiDa: 3,
  };

  it("matches by code, then exact name, then fuzzy name", () => {
    const { namiSuperTaoTraCuuThongSo } = getTestables();
    const traCuu = namiSuperTaoTraCuuThongSo(rows2026, cauHinh);

    expect(traCuu("912345", "Tên khác")).toEqual({
      thongSo: { trongLuong: 20, dai: 100, rong: 50, cao: 30, theTich: 0.15 },
      nguon: "ma",
    });
    expect(traCuu("999999", "  ghế   gỗ sồi ")?.nguon).toBe("ten");
    expect(traCuu("999999", "Tủ quần áo gỗ sồi ba cánh trắng")?.nguon).toBe(
      "ten-mo",
    );
    expect(traCuu("999999", "Kệ sách")).toBeUndefined();
  });

  it("uses 1.1 when height is empty and computes the volume", () => {
    const { namiSuperTaoTraCuuThongSo } = getTestables();
    const traCuu = namiSuperTaoTraCuuThongSo(rows2026, cauHinh);

    expect(traCuu("", "Ghế gỗ sồi")?.thongSo).toEqual({
      trongLuong: 5,
      dai: 40,
      rong: 40,
      cao: 1.1,
      theTich: 0.002,
    });
  });

  it("does not fuzzy match when disabled", () => {
    const { namiSuperTaoTraCuuThongSo } = getTestables();
    const traCuu = namiSuperTaoTraCuuThongSo(rows2026, {
      ...cauHinh,
      bat: false,
    });

    expect(traCuu("999999", "Tủ quần áo gỗ sồi ba cánh trắng")).toBeUndefined();
  });
});

describe("Upload_Nami_Super batches", () => {
  const lanNhap = (ngay, sku, hd = "JPT4001-2") => ({
    receipt_date: ngay,
    sku,
    contract_no: hd,
  });

  it("splits into batches without breaking a date|SKU|contract group", () => {
    const { namiSuperChiaDot } = getTestables();
    const items = [];
    for (let i = 0; i < 650; i++) {
      items.push(lanNhap("2026-05-25", `SKU${Math.floor(i / 3)}`));
    }

    const dots = namiSuperChiaDot(items, 300);

    expect(dots.flat()).toHaveLength(650);
    dots.forEach((dot) => expect(dot.length).toBeLessThanOrEqual(300));
    const dotCuaNhom = new Map();
    dots.forEach((dot, idx) =>
      dot.forEach((x) => {
        const key = `${x.receipt_date}|${x.sku}|${x.contract_no}`;
        expect(dotCuaNhom.get(key) ?? idx).toBe(idx);
        dotCuaNhom.set(key, idx);
      }),
    );
  });

  it("keeps a group larger than the batch size in its own batch, in order", () => {
    const { namiSuperChiaDot } = getTestables();
    const lon = Array.from({ length: 350 }, (_, i) => ({
      ...lanNhap("2026-05-25", "912345"),
      source_row: i + 2,
    }));
    const items = [lanNhap("2026-05-25", "A"), ...lon, lanNhap("2026-05-25", "B")];

    const dots = namiSuperChiaDot(items, 300);

    expect(dots.map((dot) => dot.length)).toEqual([1, 350, 1]);
    expect(dots[1].map((x) => x.source_row)).toEqual(lon.map((x) => x.source_row));
  });
});

describe("Upload_Nami_Super marking", () => {
  it("maps a receipt round to its date + quantity cells", () => {
    const { namiSuperDiaChiCapO } = getTestables();

    expect(namiSuperDiaChiCapO(120, 1)).toBe("J120:K120");
    expect(namiSuperDiaChiCapO(120, 2)).toBe("L120:M120");
    expect(namiSuperDiaChiCapO(120, 3)).toBe("N120:O120");
  });

  it("merges skipped and rejected rounds into one entry per cell pair", () => {
    const { namiSuperLapDanhSachDanhDau } = getTestables();

    const ds = namiSuperLapDanhSachDanhDau(
      [{ source_row: 7, round_number: 1, sku: "A", contract_no: "H", reason: "Thiếu SL" }],
      [
        { row: 120, round_number: 2, sku: "B", contract_no: "H", reason: "Không khớp" },
        { row: 120, round_number: 2, sku: "B", contract_no: "H", reason: "Vượt SL" },
        { row: 7, round_number: 3, sku: "A", contract_no: "H", reason: "Khác" },
      ],
    );

    expect(ds).toEqual([
      { diaChi: "J7:K7", lyDo: "Thiếu SL" },
      { diaChi: "L120:M120", lyDo: "Không khớp; Vượt SL" },
      { diaChi: "N7:O7", lyDo: "Khác" },
    ]);
  });

  it("restores a cell only while it still has the mark color", () => {
    const { namiSuperNenKhoiPhuc } = getTestables();

    expect(namiSuperNenKhoiPhuc("#F4B084", "#F4B084")).toBe(true);
    expect(namiSuperNenKhoiPhuc("f4b084", "#F4B084")).toBe(true);
    expect(namiSuperNenKhoiPhuc("#FFFF00", "#F4B084")).toBe(false);
    expect(namiSuperNenKhoiPhuc("", "#F4B084")).toBe(false);
  });

  it("reads the state sheet, skipping the header and blank rows", () => {
    const { namiSuperDocTrangThai } = getTestables();

    expect(
      namiSuperDocTrangThai([
        ["Địa chỉ", "Màu cũ", "Màu đánh dấu", "Thời điểm", "Lý do"],
        ["L120:M120", "#FFFF00", "#F4B084", "2026-05-25T08:00:00Z", "Không khớp"],
        ["", "", "", "", ""],
        ["J7:K7", "", "#F4B084", "2026-05-25T08:00:00Z", "Thiếu SL"],
      ]),
    ).toEqual([
      { diaChi: "L120:M120", mauCu: "#FFFF00", mauDanhDau: "#F4B084" },
      { diaChi: "J7:K7", mauCu: "", mauDanhDau: "#F4B084" },
    ]);
  });
});

describe("Upload_Nami_Super result merge", () => {
  it("sums counts and joins lists across batches", () => {
    const { namiSuperGopKetQua } = getTestables();

    const tong = namiSuperGopKetQua([
      {
        imported: 2,
        unchanged: 1,
        rejected: [{ row: 5, round_number: 1 }],
        receipts: ["PNK_25052026_1"],
        specs_from_sheet: [{ sku: "A", source: "ma" }],
        missing_specs: ["X"],
      },
      {
        imported: 3,
        unchanged: 0,
        rejected: [],
        receipts: ["PNK_25052026_1", "PNK_25052026_2"],
        specs_from_sheet: [{ sku: "B", source: "ten-mo" }],
        missing_specs: ["X", "Y"],
      },
    ]);

    expect(tong).toEqual({
      imported: 5,
      unchanged: 1,
      rejected: [{ row: 5, round_number: 1 }],
      receipts: ["PNK_25052026_1", "PNK_25052026_2"],
      specs_from_sheet: [
        { sku: "A", source: "ma" },
        { sku: "B", source: "ten-mo" },
      ],
      missing_specs: ["X", "Y"],
    });
  });
});
