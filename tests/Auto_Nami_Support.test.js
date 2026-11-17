import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => {
  let testables = globalThis.__AUTO_NAMI_SUPPORT_TESTABLES__;
  if (!testables) {
    throw new Error("Auto_Nami_Support test helpers were not loaded.");
  }
  return testables;
};

beforeAll(async () => {
  await import("../Auto_Nami_Support.ts");
});

describe("Auto_Nami_Support subtotal formula updates", () => {
  it("extends the last row of a SUBTOTAL formula", () => {
    let { namiSupportCapNhatCongThucSubtotal } = getTestables();

    expect(namiSupportCapNhatCongThucSubtotal("=SUBTOTAL(9, J3:J68)", 75)).toBe(
      "=SUBTOTAL(9, J3:J75)",
    );
    expect(namiSupportCapNhatCongThucSubtotal("=SUBTOTAL(9,S3:S68)", 75)).toBe(
      "=SUBTOTAL(9, S3:S75)",
    );
  });

  it("returns null when the SUBTOTAL formula already matches the last row", () => {
    let { namiSupportCapNhatCongThucSubtotal } = getTestables();

    expect(
      namiSupportCapNhatCongThucSubtotal("=SUBTOTAL(9, J3:J75)", 75),
    ).toBeNull();
  });

  it("ignores cells that are not SUBTOTAL formulas", () => {
    let { namiSupportCapNhatCongThucSubtotal } = getTestables();

    expect(namiSupportCapNhatCongThucSubtotal("TỔNG CỘNG:", 75)).toBeNull();
    expect(namiSupportCapNhatCongThucSubtotal("", 75)).toBeNull();
    expect(namiSupportCapNhatCongThucSubtotal(123, 75)).toBeNull();
    expect(namiSupportCapNhatCongThucSubtotal("=SUM(J3:J68)", 75)).toBeNull();
  });
});

describe("Auto_Nami_Support classification sorting", () => {
  const THU_TU = ["F", "A", "S", "C", "T", "M"];

  it("sorts by the configured order with KHAC last", () => {
    let { namiSupportSapXepMaPhanLoai } = getTestables();
    let input = ["KHAC", "T", "A", "M", "F"];

    let sorted = [...input].sort((a, b) =>
      namiSupportSapXepMaPhanLoai(a, b, THU_TU),
    );

    expect(sorted).toEqual(["F", "A", "T", "M", "KHAC"]);
  });

  it("puts unknown codes after configured ones (same as Auto_Nami)", () => {
    let { namiSupportSapXepMaPhanLoai } = getTestables();

    expect(namiSupportSapXepMaPhanLoai("F", "Z", THU_TU)).toBeLessThan(0);
    expect(namiSupportSapXepMaPhanLoai("Z", "KHAC", THU_TU)).toBeGreaterThan(0);
  });
});

describe("Auto_Nami_Support statistics table rebuild", () => {
  const THU_TU = ["F", "A", "S", "C", "T", "M"];

  const taoBang = (duLieu, dongCuoi = 75) => {
    let { namiSupportTaoBangThongKe } = getTestables();
    return namiSupportTaoBangThongKe(duLieu, dongCuoi, 3, "V", 68, THU_TU, 3);
  };

  it("includes newly added classification codes in both tables", () => {
    let bang = taoBang([
      { thiTruong: "JP", maPhanLoai: "T" },
      { thiTruong: "JP", maPhanLoai: "T" },
      // Mã phân loại mới xuất hiện sau khi Auto_Nami đã chạy.
      { thiTruong: "AU", maPhanLoai: "F" },
      { thiTruong: "JP", maPhanLoai: "KHAC" },
    ]);

    let dsMa = bang.thongKeMaRows.slice(2).map((row) => row[0]);
    expect(dsMa).toEqual(["F", "T", "KHAC"]);

    let dsCap = bang.dsThiTruongMaPhanLoai.map(
      (item) => `${item.thiTruong}-${item.maPhanLoai}`,
    );
    expect(dsCap).toEqual(["AU-F", "JP-T", "JP-KHAC"]);
  });

  it("builds SUMIF/SUMIFS formulas covering the current last data row", () => {
    let bang = taoBang(
      [
        { thiTruong: "JP", maPhanLoai: "T" },
        { thiTruong: "AU", maPhanLoai: "F" },
      ],
      80,
    );

    expect(bang.thongKeMaRows[2]).toEqual([
      "F",
      '=ROUND(SUMIF($E$3:$E$80,"F",$J$3:$J$80),0)',
      '=SUMIF($E$3:$E$80,"F",$N$3:$N$80)',
      '=SUMIF($E$3:$E$80,"F",$S$3:$S$80)',
    ]);

    // Bảng 1 bắt đầu hàng 3 (dưới freeze panes), 4 dòng → hàng 3–6;
    // + 1 dòng trống → bảng 2 bắt đầu hàng 8, dữ liệu từ hàng 10.
    expect(bang.thongKeMaStartRow).toBe(3);
    expect(bang.thongKeThiTruongStartRow).toBe(8);
    expect(bang.thongKeThiTruongRows[2]).toEqual([
      "AU",
      "F",
      '=ROUND(SUMIFS($J$3:$J$80,$D$3:$D$80,"AU",$E$3:$E$80,"F"),0)',
      '=SUMIFS($N$3:$N$80,$D$3:$D$80,"AU",$E$3:$E$80,"F")',
      '=SUMIFS($S$3:$S$80,$D$3:$D$80,"AU",$E$3:$E$80,"F")',
      "=Z10/68",
    ]);
  });

  it("skips blank classifications and returns null when nothing to summarize", () => {
    expect(taoBang([{ thiTruong: "JP", maPhanLoai: "" }])).toBeNull();
    expect(taoBang([])).toBeNull();
  });

  it("counts codes without market in table 1 only", () => {
    let bang = taoBang([{ thiTruong: "", maPhanLoai: "T" }]);

    expect(bang.thongKeMaRows.slice(2).map((row) => row[0])).toEqual(["T"]);
    expect(bang.dsThiTruongMaPhanLoai).toEqual([]);
  });

  it("escapes double quotes inside classification codes", () => {
    let bang = taoBang([{ thiTruong: "JP", maPhanLoai: 'A"B' }]);

    expect(bang.thongKeMaRows[2][1]).toBe(
      '=ROUND(SUMIF($E$3:$E$75,"A""B",$J$3:$J$75),0)',
    );
  });
});
