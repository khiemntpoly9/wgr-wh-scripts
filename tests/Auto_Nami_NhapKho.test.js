import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => globalThis.__AUTO_NAMI_NHAP_KHO_TESTABLES__;

beforeAll(async () => {
  await import("../Auto_Nami_NhapKho.ts");
});

describe("Auto_Nami_NhapKho helpers", () => {
  it("gộp nhiều phiếu cùng ngày thành công thức cộng", () => {
    expect(getTestables().namiNhapKhoTaoCongThuc([10, 12, 2])).toBe(
      "=10+12+2",
    );
  });

  it("gộp unique vị trí và nối liền để ghi vào cột K", () => {
    const { namiNhapKhoGopViTri, namiNhapKhoTachViTri } = getTestables();

    expect(namiNhapKhoGopViTri(["A2", "A2", "A3"])).toBe("A2A3");
    expect(namiNhapKhoGopViTri(["A2A3", "A3", "B04"])).toBe("A2A3B4");
    expect(namiNhapKhoTachViTri("A05 B44")).toEqual(["A5", "B44"]);
  });

  it("xoay vòng Nhập 2 sang Nhập 3 rồi về Nhập 1", () => {
    const { namiNhapKhoChonSlot } = getTestables();
    const empty = { dateIso: "", occupied: false };

    expect(namiNhapKhoChonSlot([empty, empty, empty])).toBe(1);
    expect(
      namiNhapKhoChonSlot([
        empty,
        { dateIso: "2026-09-14", occupied: true },
        empty,
      ]),
    ).toBe(2);
    expect(
      namiNhapKhoChonSlot([
        empty,
        { dateIso: "2026-09-14", occupied: true },
        { dateIso: "2026-09-15", occupied: true },
      ]),
    ).toBe(0);
  });

  it("không chọn slot đã có dữ liệu", () => {
    const { namiNhapKhoChonSlot } = getTestables();
    expect(
      namiNhapKhoChonSlot([
        { dateIso: "2026-09-16", occupied: true },
        { dateIso: "2026-09-14", occupied: true },
        { dateIso: "2026-09-15", occupied: true },
      ]),
    ).toBe(-1);
  });

  it("chuẩn hóa ngày Excel và ngày hiển thị", () => {
    const {
      namiNhapKhoChuanHoaNgay,
      namiNhapKhoNgayToSerial,
      namiNhapKhoTaoCongThucNgay,
    } = getTestables();
    const serial = namiNhapKhoNgayToSerial("2026-09-14");

    expect(namiNhapKhoChuanHoaNgay(serial)).toBe("2026-09-14");
    expect(namiNhapKhoChuanHoaNgay("14/09/2026")).toBe("2026-09-14");
    expect(namiNhapKhoChuanHoaNgay("2026-02-29")).toBe("");
    expect(namiNhapKhoTaoCongThucNgay("2026-09-14")).toBe(
      "=DATE(2026,9,14)",
    );
  });
});
