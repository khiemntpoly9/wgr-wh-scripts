import { beforeAll, describe, expect, it } from "vitest";

const getTestables = () => globalThis.__AUTO_NAMI_SUPPER_TESTABLES__;

beforeAll(async () => {
  await import("../Auto_Nami_Supper.ts");
});

describe("Auto_Nami_Supper configuration", () => {
  it("accepts only valid ISO dates", () => {
    const { namiSupperKiemTraNgay } = getTestables();
    expect(namiSupperKiemTraNgay("2026-09-14")).toBe(true);
    expect(namiSupperKiemTraNgay("2026-02-29")).toBe(false);
    expect(namiSupperKiemTraNgay("14/09/2026")).toBe(false);
  });

  it("requires a public HTTPS endpoint", () => {
    const { namiSupperKiemTraUrl } = getTestables();
    expect(namiSupperKiemTraUrl("https://app.wgr-wh.site/api/v1/test")).toBeNull();
    expect(namiSupperKiemTraUrl("http://app.wgr-wh.site/api/v1/test")).toContain("HTTPS");
    expect(namiSupperKiemTraUrl("https://localhost/api/test")).toContain("localhost");
  });
});
