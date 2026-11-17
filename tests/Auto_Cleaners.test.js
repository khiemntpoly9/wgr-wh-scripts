import { beforeAll, describe, expect, it } from "vitest";

const getBaymaxTestables = () => {
  let testables = globalThis.__AUTO_BAYMAX_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_BayMax test helpers were not loaded.");
  }

  return testables;
};

const getDoraemonTestables = () => {
  let testables = globalThis.__AUTO_DORAEMON_TESTABLES__;

  if (!testables) {
    throw new Error("Auto_Doraemon test helpers were not loaded.");
  }

  return testables;
};

beforeAll(async () => {
  await import("../Auto_BayMax.ts");
  await import("../Auto_Doraemon.ts");
});

describe("Shared cleaner helpers", () => {
  it("normalizes hidden spaces in Auto_BayMax", () => {
    let { lamSachBayMax } = getBaymaxTestables();

    expect(lamSachBayMax("  90 \u00A0\u200B 1019  ")).toBe("901019");
  });

  it("normalizes hidden spaces in Auto_Doraemon", () => {
    let { lamSach } = getDoraemonTestables();

    expect(lamSach("  UE \u00A0\u200B 01  ")).toBe("UE01");
  });
});
