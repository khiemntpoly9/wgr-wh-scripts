/**
 * ===================================================================
 * Sync_PrePackingList – Đồng bộ danh sách tủ tải trước từ hệ thống về sheet "pre-packing-list".
 * ===================================================================
 *
 * Cột dữ liệu (A→R, giống Auto_PXK.ts / Upload_PrePackingList.ts):
 * A STT | B Hợp đồng | C SKU | D Tên hàng | E Kích thước
 * F SL xuất đi | G ĐVT | H Đã xuất | I SL đóng gói | J ĐVT đóng gói
 * K Số lô | L Trọng lượng | M Tổng TL | N Thể tích | O Tổng TT
 * P Vị trí | Q Ghi chú | R PMC
 *
 * Luồng sử dụng:
 * 1. Mở file Excel đã dùng để upload phiếu (ô D4 có dạng "NO: 000631").
 * 2. Tạo API token tại Dashboard → API Token (Excel), điền apiToken bên dưới. Không lưu token lên repo.
 * 3. Chạy script này trong Excel Online.
 * 4. Dòng hệ thống bổ sung được chèn đúng chỗ (nền xanh lá nhạt), dòng bị sửa trên hệ thống được ghi đè
 *    (nền vàng nhạt), dòng đã xoá trên hệ thống bị xoá khỏi sheet. STT và dòng tổng được tính lại.
 *    Dòng bổ sung điền ĐVT và số lô như Auto_PXK.ts: G = "PCS", J = "PACKAGE" khi chiều cao (số thứ 3 cột E)
 *    là 1,1, ngược lại "箱/ THÙNG CARTON"; số lô trống tra sheet "lot-no" theo hợp đồng, rồi lấy số lô
 *    10 chữ số của dòng cùng hợp đồng gốc (EUF6001-1, EUF6001 3 → EUF6001). Vẫn thiếu → ô K tô đỏ.
 * 5. Lần chạy sau tự gỡ màu đánh dấu của lần trước.
 *
 * Ghép dòng Excel ↔ hệ thống theo SKU chuẩn hoá + số hợp đồng (trim, viết hoa). Dòng thiếu SKU hoặc hợp đồng
 * (dòng trống, dòng ghi chú) không bị chạm tới.
 */

const SYNC_PRE_PACKING_CONFIG = {
  tenSheet: "pre-packing-list",
  headerKeywords: ["STT", "序号"],
  listNoCell: "D4",
  /** Ghi đè số phiếu thủ công; để trống = đọc từ listNoCell. */
  listNo: "",
  apiUrl:
    "https://app.wgr-wh.site/api/v1/warehouse/pre-packing-lists/sheet-data",
  /** Để trống trên repo; người dùng tự điền token khi dán script vào Excel. */
  apiToken: "",
  markAddedColor: "#C6EFCE",
  markUpdatedColor: "#FFEB9C",
  /** Sheet tra cứu số lô: cột A hợp đồng, cột B số lô, dòng 1 là tiêu đề. */
  lotSheetName: "lot-no",
  missingLotColor: "#FF0000",
  unitColumns: {
    qtyUnitDefault: "PCS",
    packUnitCarton: "箱/ THÙNG CARTON",
    packUnitPackage: "PACKAGE",
    packageHeight: 1.1,
  },
  /** Số dòng tối đa mỗi danh sách trong tóm tắt trước khi cắt bớt. */
  summaryLimit: 50,
};

type PplSyncCell = string | number | boolean;

type PplSyncExcelRow = {
  /** Chỉ số dòng 0-based trên sheet. */
  rowIndex: number;
  /** `""` khi thiếu SKU hoặc hợp đồng. */
  key: string;
  values: PplSyncCell[];
  formulas: string[];
};

type PplSyncSheetInfo = {
  headerIndex: number;
  startIndex: number;
  /** Dòng dữ liệu cuối (có Tên hàng); `startIndex - 1` khi chưa có dòng nào. */
  lastIndex: number;
  /** `-1` khi sheet không có dòng tổng. */
  totalsIndex: number;
  totalsFormulas: string[];
  /** Cột M (12), O (14) đang là công thức ở dòng dữ liệu đầu tiên. */
  formulaCols: number[];
  /** Mọi dòng từ startIndex tới lastIndex, kể cả dòng trống. */
  rows: PplSyncExcelRow[];
};

type PplSyncSystemLine = {
  line_no: number;
  contract_no: string;
  sku: string;
  product_name: string;
  dimensions: string;
  dispatch_quantity: number;
  unit: string;
  previously_shipped_quantity: number;
  packed_quantity: number;
  pack_unit: string;
  batch_no: string;
  unit_weight: number;
  total_weight: number;
  unit_volume: number;
  total_volume: number;
  location: string;
  notes: string;
  pmc_approved: boolean;
  source: string;
};

type PplSyncCellChange = {
  sku: string;
  contractNo: string;
  column: string;
  label: string;
  oldValue: PplSyncCell;
  newValue: PplSyncCell;
};

type PplSyncMark = "added" | "updated" | null;

type PplSyncFinalRow = {
  /** Chỉ số dòng gốc trên sheet; `-1` với dòng thêm từ hệ thống. */
  rowIndex: number;
  values: PplSyncCell[];
  mark: PplSyncMark;
};

type PplSyncSummary = {
  added: { sku: string; contractNo: string; productName: string }[];
  updated: number;
  deleted: { sku: string; contractNo: string }[];
  overwritten: PplSyncCellChange[];
  /** Dòng bổ sung không tìm được số lô; điền bởi pplSyncDienDongMoi. */
  lotMissing: { sku: string; contractNo: string }[];
};

type PplSyncUnitConfig = {
  qtyUnitDefault: string;
  packUnitCarton: string;
  packUnitPackage: string;
  packageHeight: number;
};

type PplSyncPlan = {
  inserts: { afterRowIndex: number; lines: PplSyncSystemLine[] }[];
  updates: { rowIndex: number; changes: PplSyncCellChange[] }[];
  deletes: number[];
  finalRows: PplSyncFinalRow[];
  summary: PplSyncSummary;
};

const PPL_SYNC_COL_COUNT = 18;
const PPL_SYNC_FORMULA_TEMPLATES: { [col: number]: string } = {
  12: "=ROUND(L{row}*I{row}, 2)",
  14: "=ROUND(N{row}*I{row}, 2)",
};
const PPL_SYNC_SUM_RE = /^=SUM\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)$/i;
const PPL_SYNC_LOT_RE = /^\d{10}$/;

/**
 * Cột do hệ thống quản lý trên dòng đã có. Cột B, C, D, E, G, J giữ giá trị Excel.
 * - int: so sánh số nguyên; float: lệch ≤ tolerance coi là bằng (M, O bỏ qua khi ô là công thức).
 * - text: chuỗi trim; textIfSet: chỉ khi hệ thống khác rỗng; pmc: chỉ khi hệ thống true, ghi "Y".
 */
const PPL_SYNC_COMPARE_COLUMNS: {
  col: number;
  letter: string;
  label: string;
  kind: "int" | "float" | "text" | "textIfSet" | "pmc";
  field: keyof PplSyncSystemLine;
  tolerance?: number;
}[] = [
  { col: 5, letter: "F", label: "SL xuất đi", kind: "int", field: "dispatch_quantity" },
  { col: 7, letter: "H", label: "Đã xuất", kind: "int", field: "previously_shipped_quantity" },
  { col: 8, letter: "I", label: "SL đóng gói", kind: "int", field: "packed_quantity" },
  { col: 10, letter: "K", label: "Số lô", kind: "textIfSet", field: "batch_no" },
  { col: 11, letter: "L", label: "Trọng lượng", kind: "float", field: "unit_weight", tolerance: 0.0005 },
  { col: 12, letter: "M", label: "Tổng TL", kind: "float", field: "total_weight", tolerance: 0.005 },
  { col: 13, letter: "N", label: "Thể tích", kind: "float", field: "unit_volume", tolerance: 0.00005 },
  { col: 14, letter: "O", label: "Tổng TT", kind: "float", field: "total_volume", tolerance: 0.005 },
  { col: 15, letter: "P", label: "Vị trí", kind: "text", field: "location" },
  { col: 16, letter: "Q", label: "Ghi chú", kind: "textIfSet", field: "notes" },
  { col: 17, letter: "R", label: "PMC", kind: "pmc", field: "pmc_approved" },
];

const pplSyncLamSach = (giaTri: PplSyncCell | null | undefined): string =>
  String(giaTri ?? "")
    .replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]/g, "")
    .trim();

const pplSyncChuoi = (giaTri: PplSyncCell | null | undefined): string =>
  String(giaTri ?? "").trim();

/** Cùng quy tắc với prePackParseSo: chấp nhận "1.234,5" và "1,234.5". */
const pplSyncParseSo = (giaTri: PplSyncCell | null | undefined): number => {
  if (typeof giaTri === "number" && !isNaN(giaTri)) return giaTri;
  let s = pplSyncLamSach(giaTri);
  if (!s) return 0;
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (lastComma >= 0) {
    s = s.replace(",", ".");
  }
  const n = Number(s);
  return isNaN(n) ? 0 : n;
};

/** Cùng quy tắc với prePackChuanHoaSku (Upload_PrePackingList.ts) và SkuCode::clean (backend). */
const pplSyncChuanHoaSku = (sku: PplSyncCell | null | undefined): string => {
  const cleaned = pplSyncLamSach(sku);
  if (/^9\d{5,}[AB]$/i.test(cleaned)) return cleaned.toUpperCase();
  const match = cleaned.match(/^(9\d{5,})(\D.*)$/);
  return match ? match[1] : cleaned;
};

const pplSyncChuanHoaHopDong = (contractNo: PplSyncCell | null | undefined): string =>
  pplSyncChuoi(contractNo).toUpperCase();

/** Cùng quy tắc với prePackTrichXuatSoPhieu: "NO:  000463" → "000463". */
const pplSyncTrichXuatSoPhieu = (raw: PplSyncCell | null | undefined): string => {
  const s = pplSyncChuoi(raw);
  if (!s) return "";

  const noMatch = s.match(/NO\s*[：:]\s*(\S+)/i);
  if (noMatch) return noMatch[1].trim();

  const digits = s.match(/(\d[\d\s]*\d|\d+)/);
  if (digits) return digits[1].replace(/\s+/g, "");

  return s;
};

const pplSyncKiemTraApiUrl = (url: string): string | null => {
  const u = url.trim().toLowerCase();
  if (!u) return "Chưa cấu hình apiUrl.";
  if (!u.startsWith("https://")) {
    return "apiUrl phải dùng HTTPS. Office Scripts không gọi được http://.";
  }
  if (u.includes("localhost") || u.includes("127.0.0.1")) {
    return "apiUrl không được dùng localhost — Office Scripts chạy trên cloud Microsoft.";
  }
  return null;
};

const pplSyncKey = (
  sku: PplSyncCell | null | undefined,
  contractNo: PplSyncCell | null | undefined,
): string => {
  const skuKey = pplSyncChuanHoaSku(sku);
  const contractKey = pplSyncChuanHoaHopDong(contractNo);
  if (!skuKey || !contractKey) return "";
  return `${skuKey}|${contractKey}`;
};

const pplSyncDocSheet = (
  values: PplSyncCell[][],
  formulas: string[][],
  headerKeywords: string[],
): PplSyncSheetInfo | null => {
  let headerIndex = -1;
  for (let i = 0; i < values.length; i++) {
    const cell = pplSyncChuoi(values[i][0]).toUpperCase();
    if (headerKeywords.every((kw) => cell.includes(kw.toUpperCase()))) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex === -1) return null;

  const startIndex = headerIndex + 1;
  let lastIndex = startIndex - 1;
  let totalsIndex = -1;

  for (let i = startIndex; i < values.length; i++) {
    const rowFormulas = (formulas[i] ?? []).slice(0, PPL_SYNC_COL_COUNT);
    if (rowFormulas.some((f) => /^=SUM\(/i.test(pplSyncChuoi(f)))) {
      totalsIndex = i;
      break;
    }
    if (pplSyncChuoi(values[i][3]) !== "") {
      lastIndex = i;
      continue;
    }
    const rowText = values[i]
      .slice(0, PPL_SYNC_COL_COUNT)
      .map((c) => pplSyncChuoi(c))
      .join(" ")
      .toUpperCase();
    if (rowText.includes("TỔNG")) {
      totalsIndex = i;
      break;
    }
  }

  const rows: PplSyncExcelRow[] = [];
  for (let i = startIndex; i <= lastIndex; i++) {
    const rowValues: PplSyncCell[] = [];
    const rowFormulas: string[] = [];
    for (let c = 0; c < PPL_SYNC_COL_COUNT; c++) {
      rowValues.push(values[i][c] ?? "");
      rowFormulas.push(String((formulas[i] ?? [])[c] ?? ""));
    }
    rows.push({
      rowIndex: i,
      key: pplSyncChuoi(rowValues[3]) === "" ? "" : pplSyncKey(rowValues[2], rowValues[1]),
      values: rowValues,
      formulas: rowFormulas,
    });
  }

  const formulaCols: number[] = [];
  const mau = rows.find((row) => pplSyncChuoi(row.values[3]) !== "");
  if (mau) {
    [12, 14].forEach((col) => {
      if (pplSyncChuoi(mau.formulas[col]).startsWith("=")) formulaCols.push(col);
    });
  }

  const totalsFormulas: string[] = [];
  if (totalsIndex >= 0) {
    for (let c = 0; c < PPL_SYNC_COL_COUNT; c++) {
      totalsFormulas.push(String((formulas[totalsIndex] ?? [])[c] ?? ""));
    }
  }

  return {
    headerIndex,
    startIndex,
    lastIndex,
    totalsIndex,
    totalsFormulas,
    formulaCols,
    rows,
  };
};

const pplSyncCongThuc = (col: number, sheetRow: number): string =>
  PPL_SYNC_FORMULA_TEMPLATES[col].split("{row}").join(String(sheetRow));

/** 18 ô A→R của dòng mới từ hệ thống (STT điền sau). */
const pplSyncGiaTriDongMoi = (line: PplSyncSystemLine): PplSyncCell[] => [
  "",
  pplSyncChuoi(line.contract_no),
  pplSyncChuoi(line.sku),
  pplSyncChuoi(line.product_name),
  pplSyncChuoi(line.dimensions),
  Number(line.dispatch_quantity) || 0,
  pplSyncChuoi(line.unit),
  Number(line.previously_shipped_quantity) || 0,
  Number(line.packed_quantity) || 0,
  pplSyncChuoi(line.pack_unit),
  pplSyncChuoi(line.batch_no),
  Number(line.unit_weight) || 0,
  Number(line.total_weight) || 0,
  Number(line.unit_volume) || 0,
  Number(line.total_volume) || 0,
  pplSyncChuoi(line.location),
  pplSyncChuoi(line.notes),
  line.pmc_approved ? "Y" : "",
];

/** Các ô lệch giữa dòng Excel và dòng hệ thống đã ghép, theo PPL_SYNC_COMPARE_COLUMNS. */
const pplSyncSoSanhDong = (
  row: PplSyncExcelRow,
  line: PplSyncSystemLine,
): { col: number; change: PplSyncCellChange }[] => {
  const result: { col: number; change: PplSyncCellChange }[] = [];

  PPL_SYNC_COMPARE_COLUMNS.forEach((spec) => {
    const oldValue = row.values[spec.col];
    const raw = line[spec.field];
    let newValue: PplSyncCell;
    let changed: boolean;

    if (spec.kind === "int") {
      newValue = Math.round(Number(raw) || 0);
      changed = Math.round(pplSyncParseSo(oldValue)) !== newValue;
    } else if (spec.kind === "float") {
      if (pplSyncChuoi(row.formulas[spec.col]).startsWith("=")) return;
      newValue = Number(raw) || 0;
      changed = Math.abs(pplSyncParseSo(oldValue) - newValue) > (spec.tolerance ?? 0);
    } else if (spec.kind === "text") {
      newValue = pplSyncChuoi(raw as string);
      changed = pplSyncChuoi(oldValue) !== newValue;
    } else if (spec.kind === "textIfSet") {
      newValue = pplSyncChuoi(raw as string);
      changed = newValue !== "" && pplSyncChuoi(oldValue) !== newValue;
    } else {
      newValue = "Y";
      changed = raw === true && pplSyncChuoi(oldValue).toUpperCase() !== "Y";
    }

    if (!changed) return;
    result.push({
      col: spec.col,
      change: {
        sku: pplSyncChuoi(line.sku),
        contractNo: pplSyncChuoi(line.contract_no),
        column: spec.letter,
        label: spec.label,
        oldValue,
        newValue,
      },
    });
  });

  return result;
};

const pplSyncLapKeHoach = (
  info: PplSyncSheetInfo,
  systemLines: PplSyncSystemLine[],
): PplSyncPlan => {
  const queues = new Map<string, number[]>();
  systemLines.forEach((line, index) => {
    const key = pplSyncKey(line.sku, line.contract_no);
    if (!key) return;
    const queue = queues.get(key);
    if (queue) queue.push(index);
    else queues.set(key, [index]);
  });

  const matchedLine = new Map<number, number>();
  info.rows.forEach((row) => {
    if (!row.key) return;
    const queue = queues.get(row.key);
    if (queue && queue.length > 0) {
      matchedLine.set(row.rowIndex, queue.shift() as number);
    }
  });

  const used = new Set<number>(matchedLine.values());
  const deletes: number[] = [];
  const deleted: PplSyncSummary["deleted"] = [];
  const updates: PplSyncPlan["updates"] = [];
  const updatedValues = new Map<number, PplSyncCell[]>();

  info.rows.forEach((row) => {
    if (!row.key) return;
    const lineIndex = matchedLine.get(row.rowIndex);
    if (lineIndex === undefined) {
      deletes.push(row.rowIndex);
      deleted.push({ sku: pplSyncChuoi(row.values[2]), contractNo: pplSyncChuoi(row.values[1]) });
      return;
    }
    const diffs = pplSyncSoSanhDong(row, systemLines[lineIndex]);
    if (diffs.length === 0) return;
    const values = row.values.slice();
    diffs.forEach((diff) => {
      values[diff.col] = diff.change.newValue;
    });
    updatedValues.set(row.rowIndex, values);
    updates.push({ rowIndex: row.rowIndex, changes: diffs.map((diff) => diff.change) });
  });

  const deleteSet = new Set<number>(deletes);
  const survivors = info.rows.filter((row) => !deleteSet.has(row.rowIndex));

  const lastByContract = new Map<string, number>();
  survivors.forEach((row) => {
    if (row.key) lastByContract.set(pplSyncChuanHoaHopDong(row.values[1]), row.rowIndex);
  });
  const fallbackAnchor =
    survivors.length > 0 ? survivors[survivors.length - 1].rowIndex : info.headerIndex;

  const insertsByAnchor = new Map<number, PplSyncSystemLine[]>();
  systemLines.forEach((line, index) => {
    if (used.has(index)) return;
    const anchor =
      lastByContract.get(pplSyncChuanHoaHopDong(line.contract_no)) ?? fallbackAnchor;
    const group = insertsByAnchor.get(anchor);
    if (group) group.push(line);
    else insertsByAnchor.set(anchor, [line]);
  });

  const inserts = Array.from(insertsByAnchor.keys())
    .sort((a, b) => a - b)
    .map((afterRowIndex) => ({
      afterRowIndex,
      lines: insertsByAnchor.get(afterRowIndex) as PplSyncSystemLine[],
    }));

  const finalRows: PplSyncFinalRow[] = [];
  const pushInserted = (anchor: number) => {
    (insertsByAnchor.get(anchor) ?? []).forEach((line) => {
      finalRows.push({ rowIndex: -1, values: pplSyncGiaTriDongMoi(line), mark: "added" });
    });
  };

  pushInserted(info.headerIndex);
  survivors.forEach((row) => {
    const values = updatedValues.get(row.rowIndex);
    finalRows.push(
      values
        ? { rowIndex: row.rowIndex, values, mark: "updated" }
        : { rowIndex: row.rowIndex, values: row.values.slice(), mark: null },
    );
    pushInserted(row.rowIndex);
  });

  const rowByIndex = new Map<number, PplSyncExcelRow>();
  info.rows.forEach((row) => rowByIndex.set(row.rowIndex, row));

  let stt = 0;
  finalRows.forEach((finalRow, position) => {
    const original = rowByIndex.get(finalRow.rowIndex);
    if (original && !original.key) return;
    stt += 1;
    finalRow.values[0] = stt;
    if (finalRow.mark === null) return;
    const sheetRow = info.startIndex + position + 1;
    const cols = original
      ? [12, 14].filter((col) => pplSyncChuoi(original.formulas[col]).startsWith("="))
      : info.formulaCols;
    cols.forEach((col) => {
      finalRow.values[col] = pplSyncCongThuc(col, sheetRow);
    });
  });

  const summary: PplSyncSummary = {
    added: inserts.reduce(
      (all, insert) =>
        all.concat(
          insert.lines.map((line) => ({
            sku: pplSyncChuoi(line.sku),
            contractNo: pplSyncChuoi(line.contract_no),
            productName: pplSyncChuoi(line.product_name),
          })),
        ),
      [] as PplSyncSummary["added"],
    ),
    updated: updates.length,
    deleted,
    overwritten: updates.reduce(
      (all, update) => all.concat(update.changes),
      [] as PplSyncCellChange[],
    ),
    lotMissing: [],
  };

  return { inserts, updates, deletes, finalRows, summary };
};

/** Chiều cao = số thứ 3 của kích thước "a*b*c" (chấp nhận x/X và dấu phẩy thập phân). */
const pplSyncChieuCao = (dimensions: PplSyncCell): number | null => {
  const parts = pplSyncChuoi(dimensions)
    .replace(/x/gi, "*")
    .replace(/,/g, ".")
    .split("*")
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
  if (parts.length < 3) return null;
  const n = Number(parts[2]);
  return Number.isFinite(n) ? n : null;
};

/** "EUF6001-1A", "EUF6001 3" → "EUF6001". */
const pplSyncHopDongGoc = (contractNo: PplSyncCell): string => {
  const s = pplSyncChuanHoaHopDong(contractNo);
  const match = s.match(/^([A-Z0-9]+)/);
  return match ? match[1] : s;
};

const pplSyncSoLoTrong = (value: PplSyncCell): boolean => {
  const s = pplSyncChuoi(value);
  return s === "" || s === "0";
};

/** Sheet lot-no (từ A1): hợp đồng chuẩn hoá → số lô; bỏ dòng tiêu đề và dòng thiếu dữ liệu. */
const pplSyncDocBangSoLo = (values: PplSyncCell[][]): Map<string, string> => {
  const map = new Map<string, string>();
  for (let i = 1; i < values.length; i++) {
    const contract = pplSyncChuanHoaHopDong(values[i][0]);
    const lot = pplSyncChuoi(values[i][1]);
    if (contract && lot) map.set(contract, lot);
  }
  return map;
};

/**
 * Điền G, J, K cho dòng bổ sung (mark "added"). Số lô trống: lotMap theo hợp đồng, rồi số lô 10 chữ số
 * đầu tiên của dòng cùng hợp đồng gốc. Trả vị trí (trong finalRows) của dòng vẫn thiếu số lô.
 */
const pplSyncDienDongMoi = (
  plan: PplSyncPlan,
  lotMap: Map<string, string>,
  unitCfg: PplSyncUnitConfig,
): number[] => {
  const lotByBase = new Map<string, string>();
  plan.finalRows.forEach((row) => {
    const lot = pplSyncChuoi(row.values[10]);
    const base = pplSyncHopDongGoc(row.values[1]);
    if (base && PPL_SYNC_LOT_RE.test(lot) && !lotByBase.has(base)) lotByBase.set(base, lot);
  });

  const missing: number[] = [];
  plan.summary.lotMissing = [];
  plan.finalRows.forEach((row, position) => {
    if (row.mark !== "added") return;
    const height = pplSyncChieuCao(row.values[4]);
    row.values[6] = unitCfg.qtyUnitDefault;
    row.values[9] =
      height !== null && Math.abs(height - unitCfg.packageHeight) < 0.001
        ? unitCfg.packUnitPackage
        : unitCfg.packUnitCarton;

    if (!pplSyncSoLoTrong(row.values[10])) return;
    const lot =
      lotMap.get(pplSyncChuanHoaHopDong(row.values[1])) ??
      lotByBase.get(pplSyncHopDongGoc(row.values[1])) ??
      "";
    row.values[10] = lot;
    if (lot) return;
    missing.push(position);
    plan.summary.lotMissing.push({
      sku: pplSyncChuoi(row.values[2]),
      contractNo: pplSyncChuoi(row.values[1]),
    });
  });

  return missing;
};

const pplSyncKeHoachRong = (plan: PplSyncPlan): boolean =>
  plan.inserts.length === 0 && plan.updates.length === 0 && plan.deletes.length === 0;

/** Viết lại `=SUM(X{a}:X{b})` (một cột) thành `=SUM(X{startRow}:X{lastRow})`; số dòng 1-based. */
const pplSyncDongTong = (formula: string, startRow: number, lastRow: number): string => {
  const match = pplSyncChuoi(formula).match(PPL_SYNC_SUM_RE);
  if (!match || match[1].toUpperCase() !== match[3].toUpperCase()) return formula;
  if (lastRow < startRow) return "=0";
  const col = match[1].toUpperCase();
  return `=SUM(${col}${startRow}:${col}${lastRow})`;
};

/**
 * Các đoạn dòng liên tiếp có màu nền là một trong `colors`. `readColor(start, count)` trả màu chung của đoạn,
 * hoặc `""`/`null` khi đoạn có nhiều màu; khi đó chia đôi và đọc tiếp.
 */
const pplSyncTimDoanMau = (
  readColor: (start: number, count: number) => string | null,
  start: number,
  count: number,
  colors: string[],
): { start: number; count: number }[] => {
  const marks = colors.map((c) => c.toUpperCase());
  const segments: { start: number; count: number }[] = [];

  const visit = (from: number, n: number) => {
    if (n <= 0) return;
    const color = pplSyncChuoi(readColor(from, n));
    if (color !== "") {
      if (!marks.includes(color.toUpperCase())) return;
      const last = segments[segments.length - 1];
      if (last && last.start + last.count === from) last.count += n;
      else segments.push({ start: from, count: n });
      return;
    }
    if (n === 1) return;
    const half = Math.floor(n / 2);
    visit(from, half);
    visit(from + half, n - half);
  };

  visit(start, count);
  return segments;
};

const pplSyncHienThi = (value: PplSyncCell): string => {
  const s = pplSyncChuoi(value);
  return s === "" ? "(trống)" : s;
};

const pplSyncCatDanhSach = (lines: string[], limit: number): string[] => {
  if (lines.length <= limit) return lines;
  return lines.slice(0, limit).concat([`  … và ${lines.length - limit} dòng khác`]);
};

const pplSyncTomTat = (
  plan: PplSyncPlan,
  meta: { listNo: string; status: string },
  seconds: number,
  limit: number,
): string => {
  const out: string[] = [`Đồng bộ phiếu ${meta.listNo} (${meta.status}) xong trong ${seconds}s.`];

  if (meta.status !== "draft" && meta.status !== "active") {
    out.push(`Lưu ý: phiếu đang ở trạng thái ${meta.status}, dữ liệu chỉ để đối chiếu.`);
  }

  if (pplSyncKeHoachRong(plan)) {
    out.push("Đã khớp với hệ thống, không có thay đổi.");
    return out.join("\n");
  }

  const s = plan.summary;
  out.push(`Thêm ${s.added.length} · Cập nhật ${s.updated} · Xoá ${s.deleted.length}`);
  out.push("");

  if (s.added.length > 0) {
    out.push("Đã thêm (nền xanh):");
    out.push(
      ...pplSyncCatDanhSach(
        s.added.map((a) => `  + ${a.sku} · ${a.contractNo} · ${a.productName}`),
        limit,
      ),
    );
  }
  if (s.deleted.length > 0) {
    out.push("Đã xoá:");
    out.push(
      ...pplSyncCatDanhSach(
        s.deleted.map((d) => `  - ${d.sku} · ${d.contractNo}`),
        limit,
      ),
    );
  }
  if (s.overwritten.length > 0) {
    out.push("Ô đã ghi đè theo hệ thống (nền vàng):");
    out.push(
      ...pplSyncCatDanhSach(
        s.overwritten.map(
          (c) =>
            `  ~ ${c.sku} · ${c.contractNo} · ${c.label} (${c.column}): ${pplSyncHienThi(c.oldValue)} → ${pplSyncHienThi(c.newValue)}`,
        ),
        limit,
      ),
    );
  }
  if (s.lotMissing.length > 0) {
    out.push("Dòng bổ sung chưa có số lô (ô K tô đỏ):");
    out.push(
      ...pplSyncCatDanhSach(
        s.lotMissing.map((m) => `  ! ${m.sku} · ${m.contractNo}`),
        limit,
      ),
    );
  }

  out.push("");
  out.push("Chú thích: nền xanh lá nhạt = dòng bổ sung, nền vàng nhạt = dòng thay đổi.");
  out.push("Màu chỉ phản ánh lần đồng bộ này; lần chạy sau sẽ gỡ màu cũ.");

  return out.join("\n");
};

/** Gom các vị trí tăng dần thành đoạn liên tiếp. */
const pplSyncGomDoan = (positions: number[]): { start: number; count: number }[] => {
  const segments: { start: number; count: number }[] = [];
  positions.forEach((p) => {
    const last = segments[segments.length - 1];
    if (last && last.start + last.count === p) last.count += 1;
    else segments.push({ start: p, count: 1 });
  });
  return segments;
};

// @ts-ignore: Office Scripts fetch
declare function fetch(
  input: string,
  init?: {
    method?: string;
    headers?: { [key: string]: string };
    body?: string;
  },
): Promise<{
  ok: boolean;
  status: number;
  json(): Promise<unknown>;
  text(): Promise<string>;
}>;

// @ts-ignore: Office Scripts — mỗi file chạy độc lập, không trùng main()
async function main(workbook: ExcelScript.Workbook) {
  const CFG = SYNC_PRE_PACKING_CONFIG;
  const batDau = Date.now();

  const loiUrl = pplSyncKiemTraApiUrl(CFG.apiUrl);
  if (loiUrl) {
    console.log(`LỖI cấu hình: ${loiUrl}`);
    return;
  }
  if (!CFG.apiToken) {
    console.log("LỖI: Chưa cấu hình apiToken (Bearer token Sanctum).");
    return;
  }

  const sheet = workbook.getWorksheet(CFG.tenSheet);
  if (!sheet) {
    console.log(`LỖI: Không tìm thấy sheet "${CFG.tenSheet}".`);
    return;
  }

  const used = sheet.getUsedRange();
  const loiTieuDe = `LỖI: Không tìm thấy dòng tiêu đề (cột A phải chứa ${CFG.headerKeywords.join(" + ")}).`;
  if (!used) {
    console.log(loiTieuDe);
    return;
  }
  const rowCount = used.getRowIndex() + used.getRowCount();
  const readRange = sheet.getRangeByIndexes(0, 0, rowCount, PPL_SYNC_COL_COUNT);
  const info = pplSyncDocSheet(
    readRange.getValues() as PplSyncCell[][],
    readRange.getFormulas(),
    CFG.headerKeywords,
  );
  if (!info) {
    console.log(loiTieuDe);
    return;
  }

  const listNo =
    CFG.listNo.trim() || pplSyncTrichXuatSoPhieu(sheet.getRange(CFG.listNoCell).getValue());
  if (!listNo) {
    console.log(
      `LỖI: Không đọc được số phiếu từ ô ${CFG.listNoCell}. Ô phải có dạng "NO: 000631".`,
    );
    return;
  }

  let sheetData: { list_no: string; status: string; items: PplSyncSystemLine[] };
  try {
    const response = await fetch(`${CFG.apiUrl}?list_no=${encodeURIComponent(listNo)}`, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
    });
    const bodyText = await response.text();
    let body: {
      message?: string;
      errors?: Record<string, string[]>;
      data?: { list_no?: string; status?: string; items?: PplSyncSystemLine[] };
    };
    try {
      body = JSON.parse(bodyText);
    } catch {
      console.log(
        `LỖI HTTP ${response.status}: phản hồi không phải JSON hợp lệ.\n${bodyText.slice(0, 300)}`,
      );
      return;
    }

    if (response.status === 404) {
      console.log(`LỖI: Không tìm thấy danh sách tủ tải trước ${listNo} trên hệ thống.`);
      return;
    }
    if (response.status === 401 || response.status === 403) {
      console.log(
        `LỖI HTTP ${response.status}: token không hợp lệ hoặc không có quyền xem danh sách tủ tải trước.`,
      );
      return;
    }
    if (!response.ok) {
      const errorLines = [`LỖI HTTP ${response.status}: ${body.message ?? bodyText.slice(0, 300)}`];
      if (body.errors) {
        Object.keys(body.errors).forEach((field) => {
          errorLines.push(`  - ${field}: ${(body.errors as Record<string, string[]>)[field].join(", ")}`);
        });
      }
      console.log(errorLines.join("\n"));
      return;
    }
    if (!body.data || !Array.isArray(body.data.items)) {
      console.log("LỖI: Phản hồi API thiếu data.items.");
      return;
    }
    sheetData = {
      list_no: body.data.list_no ?? listNo,
      status: body.data.status ?? "",
      items: body.data.items,
    };
  } catch (err) {
    const msg = String(err);
    console.log(
      `LỖI kết nối API: ${msg} — kiểm tra apiUrl HTTPS public.\nKiểm tra apiUrl hiện tại: ${CFG.apiUrl}`,
    );
    return;
  }

  const plan = pplSyncLapKeHoach(info, sheetData.items);
  const markColors = [CFG.markAddedColor, CFG.markUpdatedColor];

  let missingLotPositions: number[] = [];
  if (plan.inserts.length > 0) {
    const lotSheet = workbook.getWorksheet(CFG.lotSheetName);
    const lotUsed = lotSheet ? lotSheet.getUsedRange() : undefined;
    const lotValues =
      lotSheet && lotUsed
        ? (lotSheet
            .getRangeByIndexes(0, 0, lotUsed.getRowIndex() + lotUsed.getRowCount(), 2)
            .getValues() as PplSyncCell[][])
        : [];
    if (!lotSheet) {
      console.log(`ℹ️ Không tìm thấy sheet "${CFG.lotSheetName}", số lô chỉ lấy theo hợp đồng gốc.`);
    }
    missingLotPositions = pplSyncDienDongMoi(plan, pplSyncDocBangSoLo(lotValues), CFG.unitColumns);
  }

  // 1. Gỡ màu đánh dấu của lần đồng bộ trước (nhận diện qua ô cột A).
  const oldCount = info.lastIndex - info.startIndex + 1;
  pplSyncTimDoanMau(
    (start, count) =>
      sheet.getRangeByIndexes(start, 0, count, 1).getFormat().getFill().getColor(),
    info.startIndex,
    oldCount,
    markColors,
  ).forEach((segment) => {
    sheet
      .getRangeByIndexes(segment.start, 0, segment.count, PPL_SYNC_COL_COUNT)
      .getFormat()
      .getFill()
      .clear();
  });

  if (!pplSyncKeHoachRong(plan)) {
    // 2. Xoá dòng từ dưới lên (cả dòng để các cột sau R không lệch hàng).
    const deletesAsc = plan.deletes.slice().sort((a, b) => a - b);
    pplSyncGomDoan(deletesAsc)
      .reverse()
      .forEach((segment) => {
        sheet
          .getRangeByIndexes(segment.start, 0, segment.count, PPL_SYNC_COL_COUNT)
          .getEntireRow()
          .delete(ExcelScript.DeleteShiftDirection.up);
      });

    // 3. Chèn dòng từ dưới lên; điểm neo là dòng gốc còn lại sau khi xoá.
    plan.inserts
      .slice()
      .sort((a, b) => b.afterRowIndex - a.afterRowIndex)
      .forEach((insert) => {
        const anchor =
          insert.afterRowIndex - deletesAsc.filter((r) => r < insert.afterRowIndex).length;
        const n = insert.lines.length;
        sheet
          .getRangeByIndexes(anchor + 1, 0, n, PPL_SYNC_COL_COUNT)
          .getEntireRow()
          .insert(ExcelScript.InsertShiftDirection.down);
        if (insert.afterRowIndex > info.headerIndex) {
          const anchorRange = sheet.getRangeByIndexes(anchor, 0, 1, PPL_SYNC_COL_COUNT);
          for (let j = 1; j <= n; j++) {
            sheet
              .getRangeByIndexes(anchor + j, 0, 1, PPL_SYNC_COL_COUNT)
              .copyFrom(anchorRange, ExcelScript.RangeCopyType.formats);
          }
        }
      });

    // 4. Ghi dòng thêm (B→R) và dòng cập nhật (F→R), gộp các dòng liên tiếp cùng loại.
    const writeGroups: { start: number; mark: PplSyncMark; rows: PplSyncCell[][] }[] = [];
    plan.finalRows.forEach((row, position) => {
      if (row.mark === null) return;
      const fromCol = row.mark === "added" ? 1 : 5;
      const cells = row.values.slice(fromCol, PPL_SYNC_COL_COUNT);
      const last = writeGroups[writeGroups.length - 1];
      if (last && last.mark === row.mark && last.start + last.rows.length === position) {
        last.rows.push(cells);
      } else {
        writeGroups.push({ start: position, mark: row.mark, rows: [cells] });
      }
    });
    writeGroups.forEach((group) => {
      const fromCol = group.mark === "added" ? 1 : 5;
      sheet
        .getRangeByIndexes(
          info.startIndex + group.start,
          fromCol,
          group.rows.length,
          PPL_SYNC_COL_COUNT - fromCol,
        )
        .setValues(group.rows);
    });

    // 5. STT trên toàn vùng dữ liệu mới.
    if (plan.finalRows.length > 0) {
      sheet
        .getRangeByIndexes(info.startIndex, 0, plan.finalRows.length, 1)
        .setValues(plan.finalRows.map((row) => [row.values[0]]));
    }

    // 6. Dòng tổng: SUM một cột theo vùng dữ liệu mới.
    if (info.totalsIndex >= 0) {
      const addedCount = plan.inserts.reduce((sum, insert) => sum + insert.lines.length, 0);
      const totalsIndex = info.totalsIndex - plan.deletes.length + addedCount;
      const firstRow = info.startIndex + 1;
      const lastRow = info.startIndex + plan.finalRows.length;
      info.totalsFormulas.forEach((formula, col) => {
        const match = pplSyncChuoi(formula).match(PPL_SYNC_SUM_RE);
        if (!match || match[1].toUpperCase() !== match[3].toUpperCase()) return;
        sheet.getCell(totalsIndex, col).setFormula(pplSyncDongTong(formula, firstRow, lastRow));
      });
    }

    // 7. Tô màu dòng thêm / dòng thay đổi.
    writeGroups.forEach((group) => {
      sheet
        .getRangeByIndexes(info.startIndex + group.start, 0, group.rows.length, PPL_SYNC_COL_COUNT)
        .getFormat()
        .getFill()
        .setColor(group.mark === "added" ? CFG.markAddedColor : CFG.markUpdatedColor);
    });
    missingLotPositions.forEach((position) => {
      sheet
        .getCell(info.startIndex + position, 10)
        .getFormat()
        .getFill()
        .setColor(CFG.missingLotColor);
    });
  }

  console.log(
    pplSyncTomTat(
      plan,
      { listNo: sheetData.list_no, status: sheetData.status },
      Math.round((Date.now() - batDau) / 1000),
      CFG.summaryLimit,
    ),
  );
}

const pplSyncTestables = {
  pplSyncKey,
  pplSyncDocSheet,
  pplSyncLapKeHoach,
  pplSyncKeHoachRong,
  pplSyncDocBangSoLo,
  pplSyncDienDongMoi,
  pplSyncDongTong,
  pplSyncTimDoanMau,
  pplSyncTomTat,
};

(
  globalThis as typeof globalThis & {
    __SYNC_PRE_PACKING_TESTABLES__?: typeof pplSyncTestables;
  }
).__SYNC_PRE_PACKING_TESTABLES__ = pplSyncTestables;
