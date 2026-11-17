import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_CONAN_TESTABLES__;
  if (!testables) {
    throw new Error("Auto_Conan test helpers were not loaded.");
  }
  return testables;
};

const ton = (tongNhap, tongXuat = 0, tonKho = 0, slDonHang = 0) => ({
  tonKho,
  tongNhap,
  tongXuat,
  slDonHang,
  soDong: 1,
});

const boLoc = (ngay = [], thang = [], nam = []) => ({ ngay, thang, nam });

beforeAll(async () => {
  await import("../Auto_Conan.ts");
});

describe("Auto_Conan", () => {
  const T = () => getTestables();
  it("tính tồn từ nhập/xuất", () => {
    let t = T();
    expect(t.conanTinhTonTuDong([10, 5, 0], [3, 2, 0])).toBe(10);
    expect(t.conanTinhTonTuDong([5], [8])).toBe(0);
  });

  it("tính tổng xuất", () => {
    expect(T().conanTinhTongXuatTuDong([3, 2, 0])).toBe(5);
  });

  it("phân tích khớp và thiếu", () => {
    let t = T();
    expect(t.conanPhanTichDong("901019A", 20, 20, 1, 0).ketQua).toBe("Khớp");
    expect(t.conanPhanTichDong("901019", 30, 10, 2, 0).thieuThung).toBe(20);
    expect(t.conanPhanTichDong("901019", 15, 0, 0, 0).ketQua).toBe(
      "Chưa nhập kho",
    );
    expect(t.conanPhanTichDong("901019", 20, 20, 1, 0).ketQua).toBe("Khớp");
  });

  it("chuẩn hóa mã hợp đồng", () => {
    expect(T().conanChuanHoaHopDong(" eu2026-001 ")).toBe("EU2026-001");
  });

  it("một dòng TP không cộng trùng K và M cùng ngày", () => {
    let t = T();
    let hangTp = new Array(15).fill("");
    hangTp[7] = "28/05/2026";
    hangTp[9] = "28/05/2026";
    hangTp[8] = 10;
    hangTp[10] = 10;
    let idxNgay = [7, 9, 11];
    let idxNhap = [8, 10, 12];
    let gop = t.conanGopNhapTheoNgayTrongDong(
      hangTp,
      idxNgay,
      idxNhap,
      "926414A",
      "AUF5003-2",
    );
    expect(gop.get("926414A|AUF5003-2|28/05/2026")).toBe(10);
  });

  it("bộ lọc ngày bật khi có ít nhất 1 tiêu chí", () => {
    let t = T();
    expect(t.conanDangBatBoLocNgay(boLoc())).toBe(false);
    expect(t.conanDangBatBoLocNgay(boLoc([13]))).toBe(true);
    expect(t.conanDangBatBoLocNgay(boLoc([], [5]))).toBe(true);
    expect(t.conanDangBatBoLocNgay(boLoc([], [], [2026]))).toBe(true);
  });

  it("khớp bộ lọc theo ngày/tháng/năm (giống Mitsuha)", () => {
    let t = T();
    expect(t.conanKhopBoLocNgay("13/05/2026", boLoc([13], [5], [2026]))).toBe(
      true,
    );
    expect(t.conanKhopBoLocNgay("14/05/2026", boLoc([13], [5], [2026]))).toBe(
      false,
    );
    expect(t.conanKhopBoLocNgay("01/06/2026", boLoc([], [6], []))).toBe(true);
    expect(t.conanKhopBoLocNgay("01/07/2026", boLoc([], [6], []))).toBe(false);
    // Bộ lọc trống → khớp mọi ngày hợp lệ.
    expect(t.conanKhopBoLocNgay("28/12/2020", boLoc())).toBe(true);
    expect(t.conanKhopBoLocNgay("không-phải-ngày", boLoc())).toBe(false);
  });

  it("tách khóa nhập ma|hopDong|ngay", () => {
    let t = T();
    expect(t.conanTachKhoaNhap("926414A|AUF5003-2|28/05/2026")).toEqual({
      ma: "926414A",
      hopDong: "AUF5003-2",
      ngay: "28/05/2026",
    });
    expect(t.conanTachKhoaNhap("rác")).toBeNull();
  });

  it("tổng nhập chi tiết theo bộ lọc + hợp đồng", () => {
    let t = T();
    let mapNgay = new Map([
      ["924125|AUF5003-2|08/06/2026", ton(50)],
      ["924125|AUF5003-2|12/06/2026", ton(10)],
      ["924125|AUF5003-2|15/06/2026", ton(5)],
    ]);
    // Bộ lọc trống → cộng tất cả ngày của HĐ.
    expect(
      t.conanTongNhapTheoBoLocHopDong(mapNgay, "924125", "AUF5003-2", boLoc())
        .tongNhap,
    ).toBe(65);
    expect(
      t.conanTongNhapTheoBoLocHopDong(mapNgay, "924125", "AUF5003-2", boLoc())
        .cacDotNhap,
    ).toEqual([
      { ngay: "08/06/2026", soLuong: 50 },
      { ngay: "12/06/2026", soLuong: 10 },
      { ngay: "15/06/2026", soLuong: 5 },
    ]);
    // Lọc theo ngày 12 → chỉ lấy lần nhập 12/06.
    expect(
      t.conanTongNhapTheoBoLocHopDong(
        mapNgay,
        "924125",
        "AUF5003-2",
        boLoc([12]),
      ).tongNhap,
    ).toBe(10);
    // Không có ngày nào khớp → 0, coDuLieu = false.
    let khongKhop = t.conanTongNhapTheoBoLocHopDong(
      mapNgay,
      "924125",
      "AUF5003-2",
      boLoc([1]),
    );
    expect(khongKhop.tongNhap).toBe(0);
    expect(khongKhop.coDuLieu).toBe(false);
    expect(khongKhop.cacDotNhap).toEqual([]);
  });

  it("tổng nhập chi tiết gộp A/B theo SKU 6 số", () => {
    let t = T();
    let mapNgay = new Map([
      ["924125A|AUF5003-2|05/06/2026", ton(4)],
      ["924125B|AUF5003-2|05/06/2026", ton(6)],
    ]);
    expect(
      t.conanTongNhapTheoBoLocHopDong(mapNgay, "924125", "AUF5003-2", boLoc())
        .tongNhap,
    ).toBe(10);
  });

  it("ngày nhập, SL nhập và SL đơn hàng theo HĐ", () => {
    let t = T();
    let mapNgay = new Map([
      ["926414A|AUF5003-2|28/05/2026", ton(10, 10)],
      ["926414|EUF5003-2|15/06/2026", ton(99, 99)],
    ]);
    let mapSku6Hd = new Map([
      ["926414|AUF5003-2", ton(10, 10, 0, 10)],
      ["926414|EUF5003-2", ton(99, 99, 0, 99)],
    ]);
    // Lọc đúng tháng 5 → chỉ AUF (28/05) được tính, EUF (15/06) = 0.
    let hang = t.conanXayDungHangVatLy(
      { sku: "926414", slNhapXuong: 10 },
      new Map([["926414", ton(109, 109, 0, 109)]]),
      new Map(),
      mapSku6Hd,
      new Map(),
      new Map(),
      new Map([["926414", ["AUF5003-2", "EUF5003-2"]]]),
      mapNgay,
      boLoc([], [5], [2026]),
      0,
    );
    let dongAu = hang.find((d) => d.hopDong === "AUF5003-2");
    let dongEu = hang.find((d) => d.hopDong === "EUF5003-2");
    expect(dongAu.tongNhapTp).toBe(10);
    expect(dongEu.tongNhapTp).toBe(0);
    expect(dongAu.cacDotNhap).toEqual([
      { ngay: "28/05/2026", soLuong: 10 },
    ]);
    expect(dongEu.cacDotNhap).toEqual([]);
    // Cột E: SL đơn hàng lấy theo từng HĐ trên Thành phẩm.
    expect(dongAu.slDonHang).toBe(10);
    expect(dongEu.slDonHang).toBe(99);
    expect(dongAu.ketQua).toBe("Khớp");
  });

  it("mở rộng nhiều hợp đồng cho một SKU", () => {
    let t = T();
    let mapNgay = new Map([
      ["901019A|EU2026-001|01/06/2026", ton(12, 12)],
      ["901019B|US2026-002|01/06/2026", ton(8, 5)],
    ]);
    let mapSku6 = new Map([["901019", ton(20, 17, 3, 20)]]);
    let mapSku6Hd = new Map([
      ["901019|EU2026-001", ton(12, 12, 0, 12)],
      ["901019|US2026-002", ton(8, 5, 3, 8)],
    ]);
    mapNgay.set("901019|EU2026-001|01/06/2026", ton(12, 12));
    mapNgay.set("901019|US2026-002|01/06/2026", ton(8, 5));
    let hang = t.conanXayDungHangVatLy(
      { sku: "901019", slNhapXuong: 20, ngayDongGoi: "15/06/2026" },
      mapSku6,
      new Map(),
      mapSku6Hd,
      new Map(),
      new Map(),
      new Map([["901019", ["EU2026-001", "US2026-002"]]]),
      mapNgay,
      boLoc(),
      0,
    );
    expect(hang).toHaveLength(2);
    expect(hang[0].ketQua).toBe("Khớp");
    expect(hang[0].tongNhapTp).toBe(12);
    expect(hang[1].tongNhapTp).toBe(8);
    expect(hang[0].slDonHang).toBe(12);
    expect(hang[1].slDonHang).toBe(8);
    expect(hang.every((d) => d.ngayDongGoi === "15/06/2026")).toBe(true);
  });

  it("so sánh tổng nhập chi tiết với SL xưởng (Thiếu)", () => {
    let t = T();
    let mapNgay = new Map([["910001|EU2026-9|03/06/2026", ton(8)]]);
    let mapSku6Hd = new Map([["910001|EU2026-9", ton(8, 0, 8, 20)]]);
    let hang = t.conanXayDungHangVatLy(
      { sku: "910001", slNhapXuong: 20 },
      new Map([["910001", ton(8, 0, 8, 20)]]),
      new Map(),
      mapSku6Hd,
      new Map(),
      new Map(),
      new Map([["910001", ["EU2026-9"]]]),
      mapNgay,
      boLoc(),
      0,
    );
    let dong = hang.find((d) => d.ketQua);
    expect(dong.ketQua).toBe("Thiếu 12 thùng");
    expect(dong.tongNhapTp).toBe(8);
    expect(dong.slDonHang).toBe(20);
  });

  it("ngày dd/mm/yyyy không đảo khi qua serial Excel", () => {
    let t = T();
    let serial = t.conanDdMmYyyySangSerial("01/06/2026");
    expect(t.conanSerialSangDdMmYyyy(serial)).toBe("01/06/2026");
    expect(t.conanNgayChuan(serial)).toBe("01/06/2026");
    let serialNgay12 = t.conanGiaTriGhiNgay("12/06/2026");
    expect(typeof serialNgay12).toBe("number");
    expect(t.conanSerialSangDdMmYyyy(serialNgay12)).toBe("12/06/2026");
    expect(t.conanGiaTriGhiNgay("08/06/2026, 12/06/2026")).toBe(
      "08/06/2026, 12/06/2026",
    );
  });

  it("chuẩn hóa SKU 6 số", () => {
    expect(T().conanChuanHoaSKU6("901019B")).toBe("901019");
  });

  it("đếm dòng có SKU cột A", () => {
    let t = T();
    let val = [["901019"], [""], ["902020"]];
    expect(t.conanDemDongCoSku(val, 2)).toEqual({ soDong: 2, hangCuoi: 4 });
  });
});
