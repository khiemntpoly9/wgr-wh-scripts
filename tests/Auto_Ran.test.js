import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_RAN_TESTABLES__;
  if (!testables) {
    throw new Error("Auto_Ran test helpers were not loaded.");
  }
  return testables;
};

beforeAll(async () => {
  await import("../Auto_Ran.ts");
});

describe("Auto_Ran date parsing helpers", () => {
  it("parses Excel serial date into yyyy-mm-dd key", () => {
    let { ranExcelSerialSangNgayKey } = getTestables();
    let serial =
      Math.floor(new Date(Date.UTC(2026, 4, 25)).getTime() / 86400000) + 25569;

    expect(ranExcelSerialSangNgayKey(serial)).toBe("2026-05-25");
  });

  it("prefers raw cell value when both text and value exist", () => {
    let { ranLayNgayKeyTuO } = getTestables();
    let serial =
      Math.floor(new Date(Date.UTC(2026, 4, 25)).getTime() / 86400000) + 25569;

    expect(ranLayNgayKeyTuO("5/25/2026", serial)).toBe("2026-05-25");
  });
});

describe("Auto_Ran report aggregation", () => {
  it("filters rows and summarizes unique tickets, quantity, weight, and volume", () => {
    let { ranTaoBaoCaoXuatCong } = getTestables();
    let headers = [
      "Ngày xuất",
      "Số phiếu",
      "STT",
      "Mã hợp đồng",
      "SKU",
      "Tên hàng",
      "Kích thước",
      "Số lượng xuất đi",
      "ĐVT",
      "Số lượng đã xuất trước đó",
      "Số lượng đóng thùng",
      "Thùng",
      "Số lô",
      "Trọng lượng",
      "Tổng trọng lượng",
      "Thể tích",
      "Tổng thể tích",
    ];
    let dataTexts = [
      headers,
      [
        "25/05/2026",
        "PXK001",
        "1",
        "JPT4001-1",
        "900001",
        "Item A",
        "40*30*20",
        "10.4",
        "PCS",
        "0.2",
        "2.1",
        "CTN",
        "LOT1",
        "5",
        "10.5",
        "0.03",
        "0.06",
      ],
      [
        "25/05/2026",
        "PXK001",
        "2",
        "JPT4001-1",
        "900002",
        "Item B",
        "40*30*20",
        "4.6",
        "PCS",
        "0.8",
        "1.2",
        "CTN",
        "LOT1",
        "3",
        "3.5",
        "0.02",
        "0.02",
      ],
      [
        "26/05/2026",
        "PXK002",
        "3",
        "JPT4001-2",
        "900003",
        "Item C",
        "40*30*20",
        "99",
        "PCS",
        "0",
        "9",
        "CTN",
        "LOT2",
        "1",
        "9",
        "0.01",
        "0.09",
      ],
    ];

    let result = ranTaoBaoCaoXuatCong(dataTexts, dataTexts, {
      targetDate: "25/05/2026",
      filterMode: "day",
      sourceStartRow: 2,
    });

    expect(result.soDongKhop).toBe(2);
    expect(result.soContHang).toBe(1);
    expect(result.tongSoLuongXuat).toBe(15);
    expect(result.tongTrongLuong).toBe(14);
    expect(result.tongTheTich).toBe(0.08);
    expect(result.resultData).toHaveLength(5);
    expect(result.resultData[0][0]).toBe("BÁO CÁO XUẤT HÀNG - ngày 25/05/2026");
    expect(result.resultData[1][1]).toBe(1);
    expect(result.resultData[1][3]).toBe(15);
    expect(result.resultData[3][0]).toBe("2026-05-25");
    expect(result.resultData[3][7]).toBe(10);
    expect(result.resultData[3][9]).toBe(0);
    expect(result.resultData[3][10]).toBe(2);
    expect(result.resultData[4][1]).toBe("PXK001");
    expect(result.resultData[4][7]).toBe(5);
    expect(result.resultData[4][9]).toBe(1);
    expect(result.resultData[4][10]).toBe(1);
    expect(result.bieuDoTheoSoPhieu).toEqual([
      {
        soPhieu: "PXK001",
        tongSoLuongXuat: 15,
        tongTrongLuong: 14,
        tongTheTich: 0.08,
      },
    ]);
  });

  it("supports week filtering", () => {
    let { ranTaoBaoCaoXuatCong } = getTestables();
    let dataTexts = [
      new Array(17).fill(""),
      ["24/05/2026", "PXK000", "", "", "", "", "", "7", "", "", "", "", "", "", "1", "", "0.1"],
      ["25/05/2026", "PXK001", "", "", "", "", "", "10", "", "", "", "", "", "", "2", "", "0.2"],
      ["31/05/2026", "PXK002", "", "", "", "", "", "20", "", "", "", "", "", "", "3", "", "0.3"],
      ["01/06/2026", "PXK003", "", "", "", "", "", "30", "", "", "", "", "", "", "4", "", "0.4"],
    ];

    let result = ranTaoBaoCaoXuatCong(dataTexts, dataTexts, {
      targetDate: "25/05/2026",
      filterMode: "week",
      sourceStartRow: 2,
    });

    expect(result.soDongKhop).toBe(2);
    expect(result.soContHang).toBe(2);
    expect(result.tongSoLuongXuat).toBe(30);
    expect(result.tongTrongLuong).toBe(5);
    expect(result.tongTheTich).toBe(0.5);
  });

  it("builds helper source data for separate charts", () => {
    let { ranTaoDuLieuNguonBieuDo } = getTestables();

    expect(
      ranTaoDuLieuNguonBieuDo([
        {
          soPhieu: "PXK001",
          tongSoLuongXuat: 15,
          tongTrongLuong: 14,
          tongTheTich: 0.08,
        },
      ]),
    ).toEqual([
      [
        "Số phiếu",
        "Tổng số lượng xuất",
        "Số phiếu",
        "Tổng trọng lượng",
        "Số phiếu",
        "Tổng thể tích",
      ],
      ["PXK001", 15, "PXK001", 14, "PXK001", 0.08],
    ]);
  });
});
