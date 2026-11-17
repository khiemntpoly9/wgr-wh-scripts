/**
 * ===================================================================
 * Auto_PXK_Super – Lấy PXK theo Container + Seal từ hệ thống.
 * ===================================================================
 *
 * Đầu vào:
 * - Sheet "sale": D7 = mã container, D8 = số seal (layout mẫu PXK).
 *   Đặt Name phạm vi sheet sale (pxk_container_no, pxk_seal_no, pxk_contracts) thì
 *   chèn / xóa hàng, cột không làm script đọc / ghi lệch ô; chưa có Name dùng D7/D8/D5.
 *
 * Kết quả:
 * - sale: chỉ đổ hàng đã lên dưới tiêu đề; không sửa header/meta mẫu
 *   (trừ D5 theo d5ContractMode); copy footer từ sheet "footer" (nếu có).
 * - not-stuffed: bảng gọn A–I (A–H như Auto_PXK, I = vị trí; không header/footer PXK).
 * - sku-loaded: mỗi lần nhập SKU một dòng (SL từng lần, không gộp tổng).
 * - tracking: theo dõi các phiếu đã lấy; chạy lại cập nhật, không trùng ticket.
 *
 * Office Scripts chạy trên Microsoft cloud: apiUrl phải là HTTPS public.
 */

const AUTO_PXK_SUPER_CONFIG = {
  sourceSheet: "sale",
  /*
   * Ô trên sheet sale: tìm theo Name trước (Formulas → Name Manager, Scope = sheet sale),
   * chưa có Name thì dùng địa chỉ fallback. Excel tự dời Name khi chèn / xóa hàng, cột.
   */
  /** Mã container (đầu vào). */
  containerNoCell: { name: "pxk_container_no", fallback: "D7" },
  /** Số seal (đầu vào). */
  sealNoCell: { name: "pxk_seal_no", fallback: "D8" },
  /** Mã hợp đồng, ghi theo d5ContractMode. */
  contractCell: { name: "pxk_contracts", fallback: "D5" },
  apiUrl:
    "https://app.wgr-wh.site/api/v1/warehouse/container-loadings/pxk-sheet-data",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909", // Bearer token Sanctum từ Dashboard → API Token (Excel)
  containerNo: "", // Để trống = đọc containerNoCell
  sealNo: "", // Để trống = đọc sealNoCell
  notStuffedSheet: "not-stuffed",
  skuLoadedSheet: "sku-loaded",
  trackingSheet: "tracking",
  fallbackHeaderRow: 10,
  headerFill: "#D9E1F2",
  compensationFill: "#FFF2CC",
  trackingHeaderFill: "#70AD47",
  fontName: "Times New Roman",
  fontSize: 14,
  packageHeight: 1.1,
  autoBalanceWeight: true,
  maxWeightAdjustPerStep: 0.5,
  minimumUnitWeight: 0.1,
  /**
   * Cách điền ô contractCell (mã hợp đồng, mặc định D5):
   * - "api": ghi chuỗi cells.D5 từ API (danh sách mã HĐ unique, cách ", ")
   * - "formula": TEXTJOIN/FILTER cột B theo S>0 như Auto_PXK
   */
  d5ContractMode: "api" as "api" | "formula",
  /**
   * Bù trừ cân: mỗi lần load lại vẫn bù.
   * Cân đích ưu tiên ô vàng M dưới dòng tổng (nếu > 0), không thì API total_gross_weight.
   * Footer được dán lại từ sheet "footer" sau khi ghi dữ liệu.
   */
};

type PxkSuperRef = { name: string; fallback: string };

/**
 * Name phạm vi sheet → Name phạm vi workbook (chỉ khi trỏ vào chính sheet này)
 * → địa chỉ fallback. Ghi tên Name không tìm thấy vào missing.
 */
const pxkSuperResolveRange = (
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  ref: PxkSuperRef,
  missing: string[],
): ExcelScript.Range => {
  const sheetItem = ref.name ? sheet.getNamedItem(ref.name) : undefined;
  if (sheetItem) return sheetItem.getRange();
  const bookItem = ref.name ? workbook.getNamedItem(ref.name) : undefined;
  if (bookItem) {
    const range = bookItem.getRange();
    if (range.getWorksheet().getName() === sheet.getName()) return range;
  }
  if (ref.name) missing.push(`${ref.name} (→ ${ref.fallback})`);
  return sheet.getRange(ref.fallback);
};

/** "'sale'!$D$5" → "D5". */
const pxkSuperLocalAddress = (range: ExcelScript.Range): string => {
  const text = String(range.getAddress() ?? "");
  const bang = text.lastIndexOf("!");
  return (bang >= 0 ? text.slice(bang + 1) : text).replace(/\$/g, "");
};

type PxkSuperItem = {
  stt?: number;
  contract_no?: string;
  sku?: string;
  product_name?: string;
  dimensions?: string;
  dispatch_quantity?: number;
  unit?: string;
  previously_shipped_quantity?: number;
  packed_quantity?: number;
  pack_unit?: string;
  batch_no?: string;
  unit_weight?: number | null;
  total_weight?: number | null;
  unit_volume?: number | null;
  total_volume?: number | null;
  notes?: string;
  location_code?: string;
  location_notes?: string;
  pmc_approved?: string;
  /** Số dòng phiếu xếp đã gộp thành 1 SKU+HĐ. */
  entry_count?: number;
  /** SL từng lần nhập (vd: [13, 15]) — sheet sku-loaded bung ra từng dòng. */
  entry_quantities?: number[];
  /** Dòng hàng được thêm để xuất bù. */
  is_compensation?: boolean;
  /** Dòng PPL đã xuất trước đó và có hàng bù tương ứng trên phiếu hiện tại. */
  is_compensation_offset?: boolean;
  /** API yêu cầu tô nền để phân biệt nghiệp vụ bù trừ. */
  highlight_compensation?: boolean;
};

type PxkSuperDroppedItem = {
  stt?: number;
  contract_no?: string;
  sku?: string;
  product_name?: string;
  dimensions?: string;
  requested_quantity?: number;
  loaded_quantity?: number;
  dropped_quantity?: number;
  unit?: string;
  batch_no?: string;
  unit_weight?: number | null;
  dropped_weight?: number | null;
  unit_volume?: number | null;
  dropped_volume?: number | null;
  notes?: string;
  location_code?: string;
};

type PxkSuperResponse = {
  message?: string;
  sheets?: { loaded?: string; dropped?: string };
  cells?: {
    D3?: string | null;
    D4?: string | null;
    D5?: string | null;
    D7?: string | null;
    D8?: string | null;
    Q6?: string | null;
  };
  ticket?: {
    id?: number;
    ticket_no?: string;
    loading_date?: string | null;
    container_no?: string;
    seal_no?: string | null;
    notes?: string | null;
    pre_packing_list?: {
      id?: number;
      list_no?: string;
      market_code?: string | null;
    } | null;
  };
  summary?: {
    total_lines?: number;
    total_quantity?: number;
    total_gross_weight?: number;
    total_volume_cbm?: number;
  };
  dropped_summary?: {
    total_lines?: number;
    total_quantity?: number;
    total_weight?: number;
    total_volume?: number;
  };
  items?: PxkSuperItem[];
  dropped_items?: PxkSuperDroppedItem[];
  candidates?: Array<{
    id: number;
    ticket_no: string;
    container_no: string;
    seal_no: string | null;
  }>;
};

const PXK_SUPER_HEADERS: (string | number | boolean)[] = [
  "STT 序号",
  "SỐ HỢP ĐỒNG 合同编号",
  "MÃ HÀNG 货号",
  "TÊN HÀNG 产品名称",
  "KÍCH THƯỚC (cm) 尺寸",
  "SỐ LƯỢNG XUẤT ĐI 出货数量",
  "ĐVT 单位",
  "ĐÃ XUẤT HÀNG 已发货数量",
  "SỐ LƯỢNG ĐÓNG GÓI 包装数量",
  "ĐVT 单位",
  "SỐ LÔ 批次号",
  "TRỌNG LƯỢNG",
  "TỔNG TRỌNG LƯỢNG (G.W./KG)总毛重",
  "THỂ TÍCH (M³) 体积",
  "TỔNG THỂ TÍCH (M³) 总体积",
  "VỊ TRÍ 仓位",
  "GHI CHÚ 备注",
  "Approve PMC y/n",
  "",
];

const PXK_SUPER_NOT_STUFFED_HEADERS: (string | number | boolean)[] = [
  "STT",
  "Mã hợp đồng",
  "SKU",
  "Tên hàng",
  "Số lượng rớt",
  "Số lượng đã xuất",
  "Tổng trọng lượng",
  "Tổng thể tích",
  "Vị trí",
];

const PXK_SUPER_SKU_LOADED_HEADERS: (string | number | boolean)[] = [
  "STT",
  "Mã hợp đồng",
  "SKU",
  "Tên hàng",
  "Số lượng đã nhập",
];

const PXK_SUPER_TRACKING_HEADERS: (string | number | boolean)[] = [
  "Ticket ID",
  "Số phiếu",
  "Ngày xếp",
  "Container",
  "Seal",
  "Pre-packing list",
  "Sheet PXK",
  "Số dòng đã nhập",
  "Tổng SL đã nhập",
  "Số dòng not-stuffed",
  "Tổng SL not-stuffed",
  "Cập nhật lúc",
  "Trạng thái",
  "Ghi chú phiếu",
];

const pxkSuperText = (value: unknown): string => String(value ?? "").trim();

const pxkSuperNumber = (value: unknown): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
};

const pxkSuperExtractCode = (raw: unknown, labels: string[]): string => {
  let value = pxkSuperText(raw);
  if (!value) return "";
  for (const label of labels) {
    value = value.replace(new RegExp(`^${label}\\s*[：:]\\s*`, "i"), "").trim();
  }
  return value.replace(/[,，。.]+$/g, "").trim();
};

const pxkSuperValidateUrl = (url: string): string | null => {
  const normalized = url.trim().toLowerCase();
  if (!normalized) return "Chưa cấu hình apiUrl.";
  if (!normalized.startsWith("https://")) return "apiUrl phải dùng HTTPS.";
  if (normalized.includes("localhost") || normalized.includes("127.0.0.1")) {
    return "apiUrl không được dùng localhost vì Office Scripts chạy trên cloud.";
  }
  return null;
};

const pxkSuperSafeSheetName = (name: string): string =>
  (name.replace(/[\\\/\?\*\[\]\:]/g, "_").trim() || "Sheet").slice(0, 31);

/** Lấy số thứ ba trong kích thước: 44*24*1,1 → 1.1. */
const pxkSuperDimensionHeight = (dimensions: unknown): number | null => {
  const parts = pxkSuperText(dimensions)
    .replace(/[×xX]/g, "*")
    .split("*")
    .map((part) => part.trim().replace(/,/g, "."));
  if (parts.length < 3) return null;
  const height = Number(parts[2].match(/-?\d+(?:\.\d+)?/)?.[0]);
  return Number.isFinite(height) ? height : null;
};

const pxkSuperPackUnit = (dimensions: unknown, packageHeight = 1.1): string => {
  const height = pxkSuperDimensionHeight(dimensions);
  return height !== null && Math.abs(height - packageHeight) < 0.000001
    ? "PACKAGE"
    : "箱/ THÙNG CARTON";
};

const pxkSuperExcelDateTime = (): number => Date.now() / 86400000 + 25569;

type PxkSuperBalanceResult = {
  weights: number[];
  total: number;
  difference: number;
  adjusted: boolean;
};

/**
 * Bù trọng lượng theo nhóm kích thước như Auto_PXK. Kết quả mô phỏng đúng
 * công thức Excel SUM(ROUND(L*I, 2)); sau khi cân theo nhóm sẽ tinh chỉnh
 * từng dòng theo bước 0,01 kg để giảm sai số làm tròn.
 */
const pxkSuperBalanceWeights = (
  dimensions: unknown[],
  unitWeights: unknown[],
  quantities: unknown[],
  targetWeight: number,
  maxAdjustPerStep = 0.5,
  minimumUnitWeight = 0.1,
): PxkSuperBalanceResult => {
  const round2 = (value: number): number => Math.round(value * 100) / 100;
  const toCents = (value: number): number => Math.round(value * 100);
  const count = Math.min(
    dimensions.length,
    unitWeights.length,
    quantities.length,
  );
  const qty = quantities.slice(0, count).map((value) => pxkSuperNumber(value));
  const sourceWeights = unitWeights
    .slice(0, count)
    .map((value) => pxkSuperNumber(value));
  const lineContribution = (weight: number, quantity: number): number =>
    round2(weight * quantity);

  if (!Number.isFinite(targetWeight) || targetWeight <= 0 || count === 0) {
    const total = round2(
      sourceWeights.reduce(
        (sum, weight, index) => sum + lineContribution(weight, qty[index]),
        0,
      ),
    );
    return {
      weights: sourceWeights,
      total,
      difference: round2(targetWeight - total),
      adjusted: false,
    };
  }

  type BalanceGroup = {
    originalWeight: number;
    weight: number;
    totalQty: number;
    indices: number[];
  };
  const groups = new Map<string, BalanceGroup>();
  for (let index = 0; index < count; index++) {
    let key = pxkSuperText(dimensions[index]).toUpperCase();
    if (!key) key = `_EMPTY_${index}`;
    if (!groups.has(key)) {
      groups.set(key, {
        originalWeight: sourceWeights[index],
        weight: sourceWeights[index],
        totalQty: 0,
        indices: [],
      });
    }
    const group = groups.get(key)!;
    group.totalQty += qty[index];
    group.indices.push(index);
  }

  let orderedGroups = Array.from(groups.values()).sort((a, b) => {
    const uniqueOrder =
      Number(a.indices.length > 1) - Number(b.indices.length > 1);
    return uniqueOrder !== 0 ? uniqueOrder : b.totalQty - a.totalQty;
  });
  let currentTotal = round2(
    orderedGroups.reduce(
      (sum, group) =>
        sum +
        group.indices.reduce(
          (groupSum, index) =>
            groupSum + lineContribution(group.weight, qty[index]),
          0,
        ),
      0,
    ),
  );
  let difference = round2(targetWeight - currentTotal);

  for (
    let iteration = 0;
    Math.abs(difference) > 0.01 && iteration < 1000;
    iteration++
  ) {
    let changed = false;
    for (const group of orderedGroups) {
      if (group.totalQty <= 0) continue;
      if (difference > 0.01) {
        const increase = Math.min(
          difference / group.totalQty,
          maxAdjustPerStep,
        );
        if (increase > 0) {
          group.weight += increase;
          difference -= increase * group.totalQty;
          changed = true;
        }
      } else if (difference < -0.01) {
        let decrease = Math.min(
          Math.abs(difference) / group.totalQty,
          maxAdjustPerStep,
        );
        decrease = Math.min(
          decrease,
          Math.max(0, group.weight - minimumUnitWeight),
        );
        if (decrease > 0) {
          group.weight -= decrease;
          difference += decrease * group.totalQty;
          changed = true;
        }
      }
      if (Math.abs(difference) <= 0.01) break;
    }
    if (!changed) break;
  }

  const balancedWeights = sourceWeights.map(() => 0);
  for (const group of orderedGroups) {
    group.weight = round2(group.weight);
    for (const index of group.indices) balancedWeights[index] = group.weight;
  }
  const simulatedTotal = (): number =>
    round2(
      balancedWeights.reduce(
        (sum, weight, index) => sum + lineContribution(weight, qty[index]),
        0,
      ),
    );
  const targetCents = toCents(targetWeight);

  // Ưu tiên giữ các SKU cùng kích thước có cùng trọng lượng.
  for (let attempt = 0; attempt < 200; attempt++) {
    const remaining = targetCents - toCents(simulatedTotal());
    if (remaining === 0) break;
    const step = remaining > 0 ? 0.01 : -0.01;
    let bestGroup: BalanceGroup | null = null;
    let bestError = Math.abs(remaining);
    for (const group of orderedGroups) {
      if (step < 0 && group.weight <= minimumUnitWeight + 0.001) continue;
      const saved = group.weight;
      group.weight = round2(saved + step);
      for (const index of group.indices) balancedWeights[index] = group.weight;
      const error = Math.abs(targetCents - toCents(simulatedTotal()));
      group.weight = saved;
      for (const index of group.indices) balancedWeights[index] = saved;
      if (error < bestError) {
        bestError = error;
        bestGroup = group;
      }
    }
    if (!bestGroup) break;
    bestGroup.weight = round2(bestGroup.weight + step);
    for (const index of bestGroup.indices)
      balancedWeights[index] = bestGroup.weight;
  }

  // Nếu làm tròn theo nhóm chưa đủ chính xác, tinh chỉnh từng dòng ít SL trước.
  const lineOrder = balancedWeights
    .map((_, index) => index)
    .sort((a, b) => qty[a] - qty[b]);
  for (let attempt = 0; attempt < 500; attempt++) {
    const remaining = targetCents - toCents(simulatedTotal());
    if (remaining === 0) break;
    const step = remaining > 0 ? 0.01 : -0.01;
    let bestIndex = -1;
    let bestError = Math.abs(remaining);
    for (const index of lineOrder) {
      if (qty[index] <= 0) continue;
      if (step < 0 && balancedWeights[index] <= minimumUnitWeight + 0.001)
        continue;
      const saved = balancedWeights[index];
      balancedWeights[index] = round2(saved + step);
      const error = Math.abs(targetCents - toCents(simulatedTotal()));
      balancedWeights[index] = saved;
      if (error < bestError) {
        bestError = error;
        bestIndex = index;
        if (error === 0) break;
      }
    }
    if (bestIndex < 0) break;
    balancedWeights[bestIndex] = round2(balancedWeights[bestIndex] + step);
  }

  currentTotal = simulatedTotal();
  difference = round2(targetWeight - currentTotal);
  const adjusted = balancedWeights.some(
    (weight, index) => Math.abs(weight - sourceWeights[index]) > 0.001,
  );
  return {
    weights: balancedWeights,
    total: currentTotal,
    difference,
    adjusted,
  };
};

const pxkSuperFindHeaderRow = (sheet: ExcelScript.Worksheet): number | null => {
  const used = sheet.getUsedRange();
  if (!used) return null;
  const lastRow = used.getRowIndex() + used.getRowCount();
  const values = sheet.getRange(`A1:A${lastRow}`).getTexts();
  for (let index = 0; index < values.length; index++) {
    const text = pxkSuperText(values[index][0]).toUpperCase();
    if (text.includes("STT") && text.includes("序号")) return index + 1;
  }
  return null;
};

const pxkSuperBorderAll = (range: ExcelScript.Range): void => {
  const format = range.getFormat();
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ].forEach((index) => {
    const border = format.getRangeBorder(index);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
    border.setColor("#000000");
  });
};

/** Lấy chuỗi nhận diện từ hàng 1 sheet mẫu "footer" (cùng workbook). */
const pxkSuperGetFooterSignature = (
  workbook: ExcelScript.Workbook,
): string | null => {
  const footerSheet = workbook.getWorksheet("footer");
  if (!footerSheet) return null;
  const fullUsedRange = footerSheet.getUsedRange();
  if (!fullUsedRange) return null;
  const footerTop = fullUsedRange.getIntersection("1:1");
  if (!footerTop) return null;
  const row = footerTop.getValues()[0];
  const parts: string[] = [];
  for (let index = 0; index < row.length; index++) {
    const text = pxkSuperText(row[index]);
    if (text) parts.push(text);
  }
  if (parts.length === 0) return null;
  let best = parts[0];
  for (let index = 1; index < parts.length; index++) {
    if (parts[index].length >= best.length) best = parts[index];
  }
  if (best.length >= 4) return best;
  const joined = parts.join(" ");
  return joined.length >= 6 ? joined : null;
};

/**
 * Dòng đầu tiên (1-based) từ fromRow có ô (A→lastCol) chứa chữ ký footer.
 * Quét nhiều cột vì cột A có thể đã bị ghi công thức STT.
 */
const pxkSuperFindFooterStartRow = (
  sheet: ExcelScript.Worksheet,
  fromRow: number,
  toRow: number,
  signature: string,
  lastColLetter = "Q",
): number | null => {
  let needle = signature.trim().toUpperCase();
  if (needle.length < 4) return null;
  if (needle.length > 40) needle = needle.substring(0, 40);
  if (toRow < fromRow) return null;
  const rows = sheet
    .getRange(`A${fromRow}:${lastColLetter}${toRow}`)
    .getValues();
  for (let index = 0; index < rows.length; index++) {
    for (let col = 0; col < rows[index].length; col++) {
      const cell = pxkSuperText(rows[index][col]).toUpperCase();
      if (cell.length > 0 && cell.includes(needle)) return fromRow + index;
    }
  }
  return null;
};

/**
 * Copy hàng 1 sheet "footer" xuống dưới dòng tổng (cách 2 dòng), giống Auto_PXK.
 * summaryRow = lastDataRow + 2 → pasteRow = summaryRow + 3.
 */
const pxkSuperPasteFooter = (
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  headerRow: number,
  rowCount: number,
): void => {
  const footerSheet = workbook.getWorksheet("footer");
  if (!footerSheet) {
    console.log(`ℹ️ Không tìm thấy sheet "footer", bỏ qua bước chèn footer.`);
    return;
  }

  const lastDataRow = headerRow + Math.max(rowCount, 0);
  const summaryRow = lastDataRow + 2;
  const pasteRow = summaryRow + 3;

  const signature = pxkSuperGetFooterSignature(workbook);
  if (signature) {
    const used = sheet.getUsedRange();
    const sheetLastRow = used
      ? used.getRowIndex() + used.getRowCount()
      : pasteRow;
    const existing = pxkSuperFindFooterStartRow(
      sheet,
      headerRow + 1,
      Math.max(sheetLastRow, pasteRow),
      signature,
    );
    if (existing !== null) {
      console.log(
        `ℹ️ Footer đã có tại dòng ${existing}, bỏ qua chèn footer trùng.`,
      );
      return;
    }
  }

  const fullUsedRange = footerSheet.getUsedRange();
  const footerRowRange = fullUsedRange
    ? fullUsedRange.getIntersection("1:1")
    : null;
  if (!footerRowRange) return;

  const targetRange = sheet
    .getCell(pasteRow - 1, 0)
    .getResizedRange(0, footerRowRange.getColumnCount() - 1);
  targetRange.copyFrom(footerRowRange, ExcelScript.RangeCopyType.all);
  console.log(
    `✅ Đã copy footer từ sheet "footer" vào dòng ${pasteRow} (cách 2 dòng từ dòng tổng).`,
  );
};

const pxkSuperClearBelowHeader = (
  sheet: ExcelScript.Worksheet,
  headerRow: number,
): void => {
  const used = sheet.getUsedRange();
  if (!used) return;
  const lastRow = used.getRowIndex() + used.getRowCount();
  if (lastRow <= headerRow) return;
  sheet.getRange(`${headerRow + 1}:${lastRow}`).setRowHidden(false);
  try {
    sheet.getRange(`P${headerRow + 1}:Q${lastRow}`).unmerge();
  } catch {
    // Không có ô gộp trong vùng.
  }
  sheet
    .getRange(`A${headerRow + 1}:S${lastRow}`)
    .clear(ExcelScript.ClearApplyTo.all);
};

/**
 * Đọc cân nặng đích đã nhập trên sheet (ô vàng dưới dòng SUM cột M),
 * trước khi clear — để lần chạy sau vẫn bù theo số user nhập.
 */
const pxkSuperFindExistingTargetWeight = (
  sheet: ExcelScript.Worksheet,
  headerRow: number,
): number => {
  const used = sheet.getUsedRange();
  if (!used) return 0;
  const lastRow = used.getRowIndex() + used.getRowCount();
  if (lastRow <= headerRow) return 0;
  sheet.getRange(`${headerRow + 1}:${lastRow}`).setRowHidden(false);
  // Đọc cả dải một lần — tránh getValue() trong vòng lặp (Office Scripts performance).
  const mRange = sheet.getRange(`M${headerRow + 1}:M${lastRow}`);
  const formulas = mRange.getFormulas();
  const values = mRange.getValues();
  for (let index = 0; index < formulas.length; index++) {
    const formula = pxkSuperText(formulas[index][0]);
    if (!/^=SUM\(/i.test(formula)) continue;
    const valueIndex = index + 1;
    if (valueIndex >= values.length) return 0;
    const value = pxkSuperNumber(values[valueIndex][0]);
    return value > 0 ? value : 0;
  }
  return 0;
};

const pxkSuperLoadedRow = (
  item: PxkSuperItem,
): (string | number | boolean)[] => [
  "",
  item.contract_no ?? "",
  item.sku ?? "",
  item.product_name ?? "",
  item.dimensions ?? "",
  pxkSuperNumber(item.dispatch_quantity),
  "PCS",
  pxkSuperNumber(item.previously_shipped_quantity),
  pxkSuperNumber(item.packed_quantity ?? item.dispatch_quantity),
  pxkSuperPackUnit(item.dimensions, AUTO_PXK_SUPER_CONFIG.packageHeight),
  item.batch_no ?? "",
  pxkSuperNumber(item.unit_weight),
  "",
  "", // N: để trống — điền công thức thể tích từ cột E (như Auto_PXK)
  "",
  pxkSuperText(item.location_code || item.location_notes), // P: Vị trí
  pxkSuperText(item.notes), // Q: Ghi chú (từ PPL, tách khỏi P)
  "", // Cột R (Approve PMC) — để trống, không điền.
  "",
];

const pxkSuperCompensationRowIndexes = (items: PxkSuperItem[]): number[] => {
  const indexes: number[] = [];
  for (let index = 0; index < items.length; index++) {
    const item = items[index];
    if (
      item.highlight_compensation === true ||
      item.is_compensation === true ||
      item.is_compensation_offset === true
    ) {
      indexes.push(index);
    }
  }
  return indexes;
};

/** Tô nền toàn dòng A–S cho hàng bù và dòng đã xuất trước đó được bù. */
const pxkSuperHighlightCompensationRows = (
  sheet: ExcelScript.Worksheet,
  headerRow: number,
  items: PxkSuperItem[],
): void => {
  const indexes = pxkSuperCompensationRowIndexes(items);
  for (const index of indexes) {
    sheet
      .getRange(`A${headerRow + index + 1}:S${headerRow + index + 1}`)
      .getFormat()
      .getFill()
      .setColor(AUTO_PXK_SUPER_CONFIG.compensationFill);
  }
  if (indexes.length > 0) {
    console.log(
      `🟨 Đã tô nền ${indexes.length} dòng hàng bù/bù trừ trên sheet ${sheet.getName()}.`,
    );
  }
};

/** not-stuffed (giống Auto_PXK): A–H + I vị trí, không layout PXK sale. */
const pxkSuperDroppedRow = (
  item: PxkSuperDroppedItem,
  stt: number,
): (string | number | boolean)[] => {
  const droppedQty = pxkSuperNumber(item.dropped_quantity);
  const unitWeight = pxkSuperNumber(item.unit_weight);
  const unitVolume = pxkSuperNumber(item.unit_volume);
  const totalWeight =
    item.dropped_weight != null && Number.isFinite(Number(item.dropped_weight))
      ? Math.round(pxkSuperNumber(item.dropped_weight) * 100) / 100
      : Math.round(droppedQty * unitWeight * 100) / 100;
  const totalVolume =
    item.dropped_volume != null && Number.isFinite(Number(item.dropped_volume))
      ? Math.round(pxkSuperNumber(item.dropped_volume) * 100) / 100
      : Math.round(droppedQty * unitVolume * 100) / 100;
  return [
    stt,
    item.contract_no ?? "",
    item.sku ?? "",
    item.product_name ?? "",
    droppedQty,
    pxkSuperNumber(item.loaded_quantity),
    totalWeight,
    totalVolume,
    pxkSuperText(item.location_code),
  ];
};

type PxkSuperSkuLoadedRow = {
  contract_no: string;
  sku: string;
  product_name: string;
  loaded_quantity: number;
};

/**
 * Mỗi lần nhập = 1 dòng (vd: 13 và 15 → 2 dòng), không gộp tổng.
 * Ưu tiên entry_quantities từ API; fallback 1 dòng theo packed/dispatch.
 */
const pxkSuperBuildSkuLoadedRows = (
  items: PxkSuperItem[],
): PxkSuperSkuLoadedRow[] => {
  const rows: PxkSuperSkuLoadedRow[] = [];

  for (const item of items) {
    const contractNo = pxkSuperText(item.contract_no);
    const sku = pxkSuperText(item.sku);
    if (!sku) continue;
    const productName = pxkSuperText(item.product_name);
    const quantities =
      Array.isArray(item.entry_quantities) && item.entry_quantities.length > 0
        ? item.entry_quantities.map((qty) => pxkSuperNumber(qty))
        : [pxkSuperNumber(item.packed_quantity ?? item.dispatch_quantity)];

    for (const qty of quantities) {
      if (qty <= 0) continue;
      rows.push({
        contract_no: contractNo,
        sku,
        product_name: productName,
        loaded_quantity: qty,
      });
    }
  }

  return rows.sort((left, right) => {
    const skuCmp = left.sku.localeCompare(right.sku, "vi", {
      sensitivity: "base",
    });
    if (skuCmp !== 0) return skuCmp;
    return left.contract_no.localeCompare(right.contract_no, "vi", {
      sensitivity: "base",
    });
  });
};

const pxkSuperSkuLoadedRow = (
  item: PxkSuperSkuLoadedRow,
  stt: number,
): (string | number | boolean)[] => [
  stt,
  item.contract_no,
  item.sku,
  item.product_name,
  item.loaded_quantity,
];

const pxkSuperApplyFormulas = (
  sheet: ExcelScript.Worksheet,
  headerRow: number,
  rowCount: number,
): void => {
  if (rowCount <= 0) return;
  const startRow = headerRow + 1;
  const lastRow = headerRow + rowCount;
  const sttFormulas: string[][] = [];
  const weightFormulas: string[][] = [];
  const volumeFormulas: string[][] = [];
  const lookupFormulas: string[][] = [];
  for (let row = startRow; row <= lastRow; row++) {
    sttFormulas.push([`=ROW()-${headerRow}`]);
    weightFormulas.push([`=ROUND(L${row}*I${row},2)`]);
    volumeFormulas.push([`=ROUND(N${row}*I${row},2)`]);
    lookupFormulas.push([`=XLOOKUP(C${row},dev!J:J,dev!K:K)`]);
  }
  sheet.getRange(`A${startRow}:A${lastRow}`).setFormulas(sttFormulas);
  sheet.getRange(`M${startRow}:M${lastRow}`).setFormulas(weightFormulas);
  sheet.getRange(`O${startRow}:O${lastRow}`).setFormulas(volumeFormulas);
  sheet.getRange(`S${startRow}:S${lastRow}`).setFormulas(lookupFormulas);

  // Cột N: công thức thể tích thùng từ kích thước E — giống Auto_PXK autoCalculateVolume.
  pxkSuperApplyVolumeFromDimensions(sheet, startRow, lastRow);
};

/**
 * Điền cột N (=ROUND((Dx*Dy*Dz)/1000000, 3)) khi ô thể tích trống,
 * lấy biểu thức từ cột E (chuẩn hóa x/X/× và dấu phẩy) — khớp Auto_PXK.
 */
const pxkSuperApplyVolumeFromDimensions = (
  sheet: ExcelScript.Worksheet,
  startRow: number,
  lastRow: number,
): void => {
  const divisor = 1000000;
  const targetValues = sheet.getRange(`N${startRow}:N${lastRow}`).getValues();
  const sourceValues = sheet.getRange(`E${startRow}:E${lastRow}`).getValues();
  let updated = 0;
  for (let index = 0; index < targetValues.length; index++) {
    const current = targetValues[index][0];
    if (current !== "" && current !== null && current !== undefined) continue;
    const raw = pxkSuperText(sourceValues[index][0]);
    if (!raw) continue;
    // Chuẩn hóa: ×/x/X → *, dấu phẩy thập phân → chấm.
    const mathExpr = raw.replace(/[×xX]/g, "*").replace(/,/g, ".");
    if (!/\d/.test(mathExpr) || mathExpr.indexOf("*") < 0) continue;
    sheet
      .getRange(`N${startRow + index}`)
      .setFormula(`=ROUND((${mathExpr})/${divisor}, 3)`);
    updated += 1;
  }
  if (updated > 0) {
    console.log(
      `✅ Đã điền công thức thể tích cột N từ kích thước E cho ${updated} dòng (=ROUND((Dx*Dy*Dz)/1000000, 3)).`,
    );
  }
};

const pxkSuperBalanceSheet = (
  sheet: ExcelScript.Worksheet,
  headerRow: number,
  rowCount: number,
  targetWeight: number,
): PxkSuperBalanceResult | null => {
  const CFG = AUTO_PXK_SUPER_CONFIG;
  if (!CFG.autoBalanceWeight || rowCount <= 0 || targetWeight <= 0) return null;
  const startRow = headerRow + 1;
  const lastRow = headerRow + rowCount;
  const result = pxkSuperBalanceWeights(
    sheet
      .getRange(`E${startRow}:E${lastRow}`)
      .getTexts()
      .map((row) => row[0]),
    sheet
      .getRange(`L${startRow}:L${lastRow}`)
      .getValues()
      .map((row) => row[0]),
    sheet
      .getRange(`I${startRow}:I${lastRow}`)
      .getValues()
      .map((row) => row[0]),
    targetWeight,
    CFG.maxWeightAdjustPerStep,
    CFG.minimumUnitWeight,
  );
  sheet
    .getRange(`L${startRow}:L${lastRow}`)
    .setValues(result.weights.map((weight) => [weight]));

  // Ô vàng ngay dưới dòng tổng (lastRow+2 = SUM → lastRow+3 = cân đích).
  // GIỮ HIỆN để lần chạy sau đọc lại; không ẩn như trước.
  const summaryRow = lastRow + 2;
  const targetRow = summaryRow + 1;
  const targetCell = sheet.getRange(`M${targetRow}`);
  const targetRounded = Math.round(targetWeight * 100) / 100;
  targetCell.setValue(targetRounded);
  targetCell.setNumberFormat("0.00"); // 9010 → 9010,00 (locale VN)
  targetCell.getFormat().getFont().setName(CFG.fontName);
  targetCell.getFormat().getFont().setSize(CFG.fontSize);
  targetCell.getFormat().getFont().setBold(true);
  targetCell.getFormat().getFill().setColor("#FFFF00");
  targetCell
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange(`${targetRow}:${targetRow}`).setRowHidden(false);

  console.log(
    result.adjusted
      ? `✅ ${sheet.getName()}: đã bù cân nặng về ${targetWeight.toFixed(2)} kg; sai số còn lại ${result.difference.toFixed(2)} kg.`
      : `ℹ️ ${sheet.getName()}: tổng cân nặng đã khớp mục tiêu ${targetWeight.toFixed(2)} kg.`,
  );
  return result;
};

const pxkSuperFormatSheet = (
  sheet: ExcelScript.Worksheet,
  headerRow: number,
  rowCount: number,
  headerFill: string,
  fontName: string,
  fontSize: number,
  touchHeader = true,
): void => {
  const lastDataRow = headerRow + rowCount;
  const startRow = headerRow + 1;

  if (touchHeader) {
    const finalRow = Math.max(headerRow, lastDataRow);
    const full = sheet.getRange(`A${headerRow}:S${finalRow}`);
    sheet.getRange(`A1:S${finalRow}`).getFormat().getFont().setName(fontName);
    sheet.getRange(`A1:S${finalRow}`).getFormat().getFont().setSize(fontSize);
    full.getFormat().getFont().setName(fontName);
    full.getFormat().getFont().setSize(fontSize);
    full.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    pxkSuperBorderAll(full);

    const header = sheet.getRange(`A${headerRow}:S${headerRow}`);
    header.getFormat().getFont().setBold(true);
    header.getFormat().getFill().setColor(headerFill);
    header
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    header.getFormat().setWrapText(true);
    header.getFormat().setRowHeight(42);
  }

  if (rowCount > 0) {
    const data = sheet.getRange(`A${startRow}:S${lastDataRow}`);
    data.getFormat().getFont().setName(fontName);
    data.getFormat().getFont().setSize(fontSize);
    data.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    pxkSuperBorderAll(data);

    sheet
      .getRange(`A${startRow}:C${lastDataRow}`)
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    sheet
      .getRange(`D${startRow}:D${lastDataRow}`)
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet
      .getRange(`E${startRow}:O${lastDataRow}`)
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    sheet
      .getRange(`P${startRow}:Q${lastDataRow}`)
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    sheet
      .getRange(`R${startRow}:S${lastDataRow}`)
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    sheet.getRange(`F${startRow}:I${lastDataRow}`).setNumberFormat("0");
    sheet.getRange(`L${startRow}:M${lastDataRow}`).setNumberFormat("0.00");
    sheet.getRange(`N${startRow}:O${lastDataRow}`).setNumberFormat("0.00");

    const summaryRow = lastDataRow + 2;
    ["F", "I", "M", "O"].forEach((column) => {
      const cell = sheet.getRange(`${column}${summaryRow}`);
      cell.setFormula(`=SUM(${column}${startRow}:${column}${lastDataRow})`);
      cell.getFormat().getFont().setBold(true);
      cell.getFormat().getFont().setName(fontName);
      cell.getFormat().getFont().setSize(fontSize);
      cell.getFormat().getFill().setColor("#FFFF00");
      cell
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      // M/O luôn hiện 2 số thập phân (9010 → 9010,00 theo locale).
      if (column === "M" || column === "O") {
        cell.setNumberFormat("0.00");
      } else {
        cell.setNumberFormat("0");
      }
      pxkSuperBorderAll(cell);
    });
  }

  if (touchHeader) {
    sheet.getRange("H:H").setColumnHidden(true);
    sheet.getRange("S:S").setColumnHidden(true);
    sheet.getRange("D:D").getFormat().setWrapText(false);
    // Giữ tỷ lệ gần mẫu Auto_PXK, tránh autofit tiêu đề song ngữ thành cột quá rộng.
    sheet.getRange("A:A").getFormat().setColumnWidth(45);
    sheet.getRange("B:B").getFormat().setColumnWidth(115);
    sheet.getRange("C:C").getFormat().setColumnWidth(85);
    sheet.getRange("D:D").getFormat().setColumnWidth(220);
    sheet.getRange("E:E").getFormat().setColumnWidth(105);
    sheet.getRange("F:K").getFormat().setColumnWidth(85);
    sheet.getRange("L:O").getFormat().setColumnWidth(95);
    sheet.getRange("P:Q").getFormat().setColumnWidth(110);
    sheet.getRange("R:R").getFormat().setColumnWidth(85);
    sheet.getFreezePanes().unfreeze();
    sheet.getFreezePanes().freezeRows(headerRow);
  }
};

/**
 * Đổ dữ liệu vào sheet sale: chỉ ghi dưới hàng tiêu đề.
 * Không đụng meta/header (D3/D4/D7/D8, font, fill, cột, freeze...),
 * trừ ô mã hợp đồng (contractCell) theo d5ContractMode.
 *
 * Cân đích bù trừ: ưu tiên ô vàng M dưới dòng tổng (nếu user đã nhập),
 * không có thì dùng summary.total_gross_weight từ API. Footer dán lại sau cùng.
 */
const pxkSuperFillSaleSheet = (
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  headerRow: number,
  rows: (string | number | boolean)[][],
  items: PxkSuperItem[],
  apiTargetWeight: number,
  apiContracts: string,
  contractRange: ExcelScript.Range,
): void => {
  const CFG = AUTO_PXK_SUPER_CONFIG;
  const sheetTarget = pxkSuperFindExistingTargetWeight(sheet, headerRow);
  const targetWeight = sheetTarget > 0 ? sheetTarget : apiTargetWeight;
  console.log(
    sheetTarget > 0
      ? `⚖️ Cân đích từ ô vàng trên sheet: ${targetWeight} kg (giữ qua lần load lại).`
      : `⚖️ Cân đích từ API: ${targetWeight} kg (có thể sửa ô vàng M dưới dòng tổng rồi chạy lại).`,
  );

  pxkSuperClearBelowHeader(sheet, headerRow);
  if (rows.length > 0) {
    sheet
      .getRangeByIndexes(headerRow, 0, rows.length, PXK_SUPER_HEADERS.length)
      .setValues(rows);
  }
  pxkSuperApplyFormulas(sheet, headerRow, rows.length);
  pxkSuperFillD5Contracts(contractRange, headerRow, rows.length, apiContracts);
  // Summary trước, rồi bù trừ + ghi lại ô vàng, cuối cùng footer.
  pxkSuperFormatSheet(
    sheet,
    headerRow,
    rows.length,
    CFG.headerFill,
    CFG.fontName,
    CFG.fontSize,
    false,
  );
  pxkSuperHighlightCompensationRows(sheet, headerRow, items);
  pxkSuperBalanceSheet(sheet, headerRow, rows.length, targetWeight);
  pxkSuperPasteFooter(workbook, sheet, headerRow, rows.length);
};

/** Ghép mã HĐ unique từ items (fallback khi API chưa có cells.D5). */
const pxkSuperJoinContractsFromItems = (items: PxkSuperItem[]): string => {
  const ordered: string[] = [];
  for (let index = 0; index < items.length; index++) {
    const no = pxkSuperText(items[index].contract_no);
    if (!no) continue;
    const key = no.toUpperCase();
    let exists = false;
    for (let prev = 0; prev < ordered.length; prev++) {
      if (ordered[prev].toUpperCase() === key) {
        exists = true;
        break;
      }
    }
    if (!exists) ordered.push(no);
  }
  return ordered.join(", ");
};

/**
 * Ô mã hợp đồng (contractCell, mặc định D5) — 2 chế độ (d5ContractMode):
 * - api: ghi chuỗi từ API (cells.D5)
 * - formula: TEXTJOIN/FILTER cột B theo S>0 như Auto_PXK
 */
const pxkSuperFillD5Contracts = (
  contractRange: ExcelScript.Range,
  headerRow: number,
  rowCount: number,
  apiContracts: string,
): void => {
  const mode = AUTO_PXK_SUPER_CONFIG.d5ContractMode;
  const address = pxkSuperLocalAddress(contractRange);
  if (mode === "api") {
    contractRange.setValue(apiContracts);
    console.log(
      apiContracts
        ? `✅ ${address} (api): ${apiContracts}`
        : `ℹ️ ${address} (api): không có mã hợp đồng.`,
    );
    return;
  }

  if (rowCount <= 0) {
    contractRange.setValue("");
    return;
  }
  const startRow = headerRow + 1;
  const lastRow = headerRow + rowCount;
  contractRange.setFormula(
    `=TEXTJOIN(", ",TRUE,UNIQUE(FILTER(B${startRow}:B${lastRow},IFERROR(S${startRow}:S${lastRow}>0,FALSE))))`,
  );
  console.log(
    `✅ ${address} (formula): TEXTJOIN B${startRow}:B${lastRow} / S${startRow}:S${lastRow}.`,
  );
};

/**
 * Sheet not-stuffed giống Auto_PXK:
 * - Dòng 1: SUBTOTAL(E/G/H)
 * - Dòng 2: tiêu đề bảng A–I (I = vị trí)
 * - Dòng 3+: dữ liệu
 * Không copy header/footer PXK sale.
 */
const pxkSuperWriteNotStuffedSheet = (
  workbook: ExcelScript.Workbook,
  items: PxkSuperDroppedItem[],
): void => {
  const CFG = AUTO_PXK_SUPER_CONFIG;
  const safeName = pxkSuperSafeSheetName(CFG.notStuffedSheet);
  let sheet = workbook.getWorksheet(safeName);

  if (items.length === 0) {
    if (sheet) {
      sheet.setVisibility(ExcelScript.SheetVisibility.visible);
      const oldRange = sheet.getUsedRange();
      if (oldRange) oldRange.clear(ExcelScript.ClearApplyTo.all);
      console.log(
        `ℹ️ Không có sản phẩm rớt — đã xóa dữ liệu cũ trên sheet "${safeName}".`,
      );
    } else {
      console.log(
        `ℹ️ Không có sản phẩm nào bị rớt, không cần tạo/cập nhật sheet "${safeName}".`,
      );
    }
    return;
  }

  if (!sheet) {
    sheet = workbook.addWorksheet(safeName);
  } else {
    sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    const oldRange = sheet.getUsedRange();
    if (oldRange) oldRange.clear(ExcelScript.ClearApplyTo.all);
  }

  // getWorksheet có thể undefined theo kiểu; sau addWorksheet chắc chắn có sheet.
  const nsSheet = workbook.getWorksheet(safeName);
  if (!nsSheet) return;

  const rows = items.map((item, index) => pxkSuperDroppedRow(item, index + 1));
  const lastDataRow = 2 + rows.length;

  nsSheet.getRange("A2:I2").setValues([PXK_SUPER_NOT_STUFFED_HEADERS]);
  const headerRange = nsSheet.getRange("A2:I2");
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setName(CFG.fontName);
  headerRange.getFormat().getFont().setSize(CFG.fontSize);
  headerRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRange.getFormat().getFill().setColor(CFG.headerFill);

  const dataRange = nsSheet.getRange(`A3:I${lastDataRow}`);
  dataRange.setValues(rows);
  dataRange.getFormat().getFont().setName(CFG.fontName);
  dataRange.getFormat().getFont().setSize(CFG.fontSize);
  dataRange
    .getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  nsSheet
    .getRange(`A3:C${lastDataRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  nsSheet
    .getRange(`D3:D${lastDataRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  nsSheet
    .getRange(`E3:I${lastDataRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  pxkSuperBorderAll(nsSheet.getRange(`A2:I${lastDataRow}`));

  ["E", "G", "H"].forEach((col) => {
    const cell = nsSheet.getRange(`${col}1`);
    cell.setFormula(`=SUBTOTAL(9, ${col}3:${col}${lastDataRow})`);
    const font = cell.getFormat().getFont();
    font.setName(CFG.fontName);
    font.setSize(CFG.fontSize);
    font.setBold(true);
    font.setColor("#FF0000");
    cell
      .getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  });

  nsSheet.getFreezePanes().freezeRows(2);
  nsSheet.getRange("A:I").getFormat().autofitColumns();
  console.log(
    `✅ Đã ghi ${rows.length} sản phẩm rớt sang sheet "${safeName}".`,
  );
};

/**
 * Sheet sku-loaded: mỗi lần nhập SKU = 1 dòng (không gộp tổng).
 * Cột: STT | Mã HĐ | SKU | Tên hàng | Số lượng đã nhập.
 */
const pxkSuperWriteSkuLoadedSheet = (
  workbook: ExcelScript.Workbook,
  items: PxkSuperItem[],
): void => {
  const CFG = AUTO_PXK_SUPER_CONFIG;
  const safeName = pxkSuperSafeSheetName(CFG.skuLoadedSheet);
  let sheet = workbook.getWorksheet(safeName);
  const summary = pxkSuperBuildSkuLoadedRows(items);

  if (summary.length === 0) {
    if (sheet) {
      sheet.setVisibility(ExcelScript.SheetVisibility.visible);
      const oldRange = sheet.getUsedRange();
      if (oldRange) oldRange.clear(ExcelScript.ClearApplyTo.all);
      console.log(
        `ℹ️ Không có SKU đã nhập — đã xóa dữ liệu cũ trên sheet "${safeName}".`,
      );
    } else {
      console.log(
        `ℹ️ Không có SKU đã nhập, không cần tạo/cập nhật sheet "${safeName}".`,
      );
    }
    return;
  }

  if (!sheet) {
    sheet = workbook.addWorksheet(safeName);
  } else {
    sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    const oldRange = sheet.getUsedRange();
    if (oldRange) oldRange.clear(ExcelScript.ClearApplyTo.all);
  }

  const skuSheet = workbook.getWorksheet(safeName);
  if (!skuSheet) return;

  const rows = summary.map((item, index) =>
    pxkSuperSkuLoadedRow(item, index + 1),
  );
  const lastDataRow = 2 + rows.length;

  skuSheet.getRange("A2:E2").setValues([PXK_SUPER_SKU_LOADED_HEADERS]);
  const headerRange = skuSheet.getRange("A2:E2");
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setName(CFG.fontName);
  headerRange.getFormat().getFont().setSize(CFG.fontSize);
  headerRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRange.getFormat().getFill().setColor(CFG.headerFill);

  const dataRange = skuSheet.getRange(`A3:E${lastDataRow}`);
  dataRange.setValues(rows);
  dataRange.getFormat().getFont().setName(CFG.fontName);
  dataRange.getFormat().getFont().setSize(CFG.fontSize);
  dataRange
    .getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  skuSheet
    .getRange(`A3:C${lastDataRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  skuSheet
    .getRange(`D3:D${lastDataRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  skuSheet
    .getRange(`E3:E${lastDataRow}`)
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  skuSheet.getRange(`E3:E${lastDataRow}`).setNumberFormat("0");

  pxkSuperBorderAll(skuSheet.getRange(`A2:E${lastDataRow}`));

  const totalCell = skuSheet.getRange("E1");
  totalCell.setFormula(`=SUBTOTAL(9, E3:E${lastDataRow})`);
  const font = totalCell.getFormat().getFont();
  font.setName(CFG.fontName);
  font.setSize(CFG.fontSize);
  font.setBold(true);
  font.setColor("#FF0000");
  totalCell
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  skuSheet.getFreezePanes().freezeRows(2);
  skuSheet.getRange("A:E").getFormat().autofitColumns();
  console.log(
    `✅ Đã ghi ${rows.length} lần nhập SKU sang sheet "${safeName}".`,
  );
};

const pxkSuperTrackingKey = (
  ticketId: unknown,
  ticketNo: unknown,
  containerNo: unknown,
  sealNo: unknown,
): string => {
  const id = pxkSuperText(ticketId);
  return id
    ? `ID:${id}`
    : `${pxkSuperText(ticketNo).toUpperCase()}|${pxkSuperText(containerNo).toUpperCase()}|${pxkSuperText(sealNo).toUpperCase()}`;
};

const pxkSuperUpdateTracking = (
  workbook: ExcelScript.Workbook,
  body: PxkSuperResponse,
  loadedSheetName: string,
  loadedCount: number,
  loadedQty: number,
  droppedCount: number,
  droppedQty: number,
): void => {
  const CFG = AUTO_PXK_SUPER_CONFIG;
  let sheet = workbook.getWorksheet(CFG.trackingSheet);
  if (!sheet) sheet = workbook.addWorksheet(CFG.trackingSheet);
  sheet.setVisibility(ExcelScript.SheetVisibility.visible);

  const used = sheet.getUsedRange();
  if (!used || pxkSuperText(sheet.getRange("A1").getValue()) !== "Ticket ID") {
    if (used) used.clear(ExcelScript.ClearApplyTo.all);
  }
  sheet.getRange("A1:N1").setValues([PXK_SUPER_TRACKING_HEADERS]);

  const ticketId = body.ticket?.id ?? "";
  const ticketNo = body.ticket?.ticket_no ?? body.cells?.D4 ?? "";
  const containerNo = body.ticket?.container_no ?? body.cells?.D7 ?? "";
  const sealNo = body.ticket?.seal_no ?? body.cells?.D8 ?? "";
  const wantedKey = pxkSuperTrackingKey(
    ticketId,
    ticketNo,
    containerNo,
    sealNo,
  );
  const trackingUsed = sheet.getUsedRange();
  const lastRow = trackingUsed
    ? trackingUsed.getRowIndex() + trackingUsed.getRowCount()
    : 1;
  let targetRow = lastRow + 1;
  if (lastRow >= 2) {
    const old = sheet.getRange(`A2:E${lastRow}`).getValues();
    for (let index = 0; index < old.length; index++) {
      if (
        pxkSuperTrackingKey(
          old[index][0],
          old[index][1],
          old[index][3],
          old[index][4],
        ) === wantedKey
      ) {
        targetRow = index + 2;
        break;
      }
    }
  }

  sheet
    .getRange(`A${targetRow}:N${targetRow}`)
    .setValues([
      [
        ticketId,
        ticketNo,
        body.ticket?.loading_date ?? body.cells?.D3 ?? "",
        containerNo,
        sealNo,
        body.ticket?.pre_packing_list?.list_no ?? "",
        loadedSheetName,
        loadedCount,
        loadedQty,
        droppedCount,
        droppedQty,
        pxkSuperExcelDateTime(),
        "ĐÃ LẤY TỪ HỆ THỐNG",
        pxkSuperText(body.ticket?.notes),
      ],
    ]);

  const finalLastRow = Math.max(lastRow, targetRow);
  const header = sheet.getRange("A1:N1");
  header.getFormat().getFont().setBold(true);
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFill().setColor(CFG.trackingHeaderFill);
  header
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet
    .getRange(`A1:N${finalLastRow}`)
    .getFormat()
    .getFont()
    .setName(CFG.fontName);
  sheet
    .getRange(`A1:N${finalLastRow}`)
    .getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange(`L2:L${finalLastRow}`).setNumberFormat("dd/mm/yyyy hh:mm");
  sheet.getRange(`H2:K${finalLastRow}`).setNumberFormat("0");
  pxkSuperBorderAll(sheet.getRange(`A1:N${finalLastRow}`));
  sheet.getRange(`A1:M${finalLastRow}`).getFormat().autofitColumns();
  sheet.getRange("G:G").getFormat().setColumnWidth(150);
  sheet.getRange("M:M").getFormat().setColumnWidth(150);
  sheet.getRange("N:N").getFormat().setColumnWidth(300);
  sheet.getRange(`N2:N${finalLastRow}`).getFormat().setWrapText(true);
  sheet.getFreezePanes().unfreeze();
  sheet.getFreezePanes().freezeRows(1);
};

// @ts-ignore: Office Scripts cung cấp fetch ở runtime.
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
  text(): Promise<string>;
}>;

// @ts-ignore: Mỗi Office Script chạy độc lập nên có thể dùng tên main.
async function main(workbook: ExcelScript.Workbook) {
  const CFG = AUTO_PXK_SUPER_CONFIG;
  const urlError = pxkSuperValidateUrl(CFG.apiUrl);
  if (urlError) {
    console.log(`LỖI cấu hình: ${urlError}`);
    return;
  }
  if (!CFG.apiToken.trim()) {
    console.log(
      "LỖI: Chưa cấu hình apiToken. Hãy tạo Bearer token tại Dashboard → API Token (Excel).",
    );
    return;
  }

  const source = workbook.getWorksheet(CFG.sourceSheet);
  if (!source) {
    console.log(`LỖI: Không tìm thấy sheet nguồn \"${CFG.sourceSheet}\".`);
    return;
  }

  const missingNames: string[] = [];
  const containerRange = pxkSuperResolveRange(
    workbook,
    source,
    CFG.containerNoCell,
    missingNames,
  );
  const sealRange = pxkSuperResolveRange(
    workbook,
    source,
    CFG.sealNoCell,
    missingNames,
  );
  const contractRange = pxkSuperResolveRange(
    workbook,
    source,
    CFG.contractCell,
    missingNames,
  );

  const containerNo =
    CFG.containerNo.trim() ||
    pxkSuperExtractCode(containerRange.getValue(), [
      "CONT",
      "CONTAINER",
      "MÃ CONT",
      "MA CONT",
    ]);
  const sealNo =
    CFG.sealNo.trim() ||
    pxkSuperExtractCode(sealRange.getValue(), [
      "SEAL",
      "CHÌ",
      "CHI",
      "SỐ SEAL",
      "SO SEAL",
    ]);
  if (!containerNo || !sealNo) {
    console.log(
      `LỖI: Thiếu ${!containerNo ? `mã container tại ${pxkSuperLocalAddress(containerRange)}` : `số seal tại ${pxkSuperLocalAddress(sealRange)}`}.`,
    );
    return;
  }

  const url =
    CFG.apiUrl.replace(/\?.*$/, "") +
    `?container_no=${encodeURIComponent(containerNo)}&seal_no=${encodeURIComponent(sealNo)}`;
  console.log(
    `Đang lấy PXK từ hệ thống: Cont=${containerNo} | Seal=${sealNo}...`,
  );

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${CFG.apiToken}`,
      },
    });
    const bodyText = await response.text();
    let body: PxkSuperResponse;
    try {
      body = JSON.parse(bodyText) as PxkSuperResponse;
    } catch {
      console.log(`LỖI: API trả dữ liệu không phải JSON (${response.status}).`);
      return;
    }

    if (response.status === 409) {
      const candidates = (body.candidates ?? [])
        .map(
          (item) =>
            `- ${item.ticket_no} | ${item.container_no} | ${item.seal_no ?? "—"} | ID ${item.id}`,
        )
        .join("\n");
      console.log(
        `${body.message ?? "Có nhiều phiếu trùng cont + seal."}\n${candidates}`,
      );
      return;
    }
    if (!response.ok) {
      console.log(`LỖI API ${response.status}: ${body.message ?? bodyText}`);
      return;
    }

    const loadedItems = body.items ?? [];
    const droppedItems = body.dropped_items ?? [];
    const saleHeaderRow =
      pxkSuperFindHeaderRow(source) ?? CFG.fallbackHeaderRow;
    const loadedRows = loadedItems.map((item) => pxkSuperLoadedRow(item));
    const targetWeight = pxkSuperNumber(body.summary?.total_gross_weight);
    const apiContracts =
      pxkSuperText(body.cells?.D5) ||
      pxkSuperJoinContractsFromItems(loadedItems);

    // Đổ hàng đã lên vào sale — không đụng header/meta mẫu (trừ ô mã hợp đồng).
    pxkSuperFillSaleSheet(
      workbook,
      source,
      saleHeaderRow,
      loadedRows,
      loadedItems,
      targetWeight,
      apiContracts,
      contractRange,
    );
    pxkSuperWriteNotStuffedSheet(workbook, droppedItems);
    pxkSuperWriteSkuLoadedSheet(workbook, loadedItems);

    const loadedQty =
      body.summary?.total_quantity ??
      loadedItems.reduce(
        (sum, item) =>
          sum + pxkSuperNumber(item.packed_quantity ?? item.dispatch_quantity),
        0,
      );
    const droppedQty =
      body.dropped_summary?.total_quantity ??
      droppedItems.reduce(
        (sum, item) => sum + pxkSuperNumber(item.dropped_quantity),
        0,
      );
    pxkSuperUpdateTracking(
      workbook,
      body,
      source.getName(),
      loadedItems.length,
      loadedQty,
      droppedItems.length,
      droppedQty,
    );

    source.activate();
    console.log("=========================================");
    console.log("  HOÀN TẤT AUTO_PXK_SUPER");
    console.log("=========================================");
    console.log(
      `• ${source.getName()}: ${loadedItems.length} dòng, ${loadedQty} PCS`,
    );
    console.log(
      `• ${CFG.notStuffedSheet}: ${droppedItems.length} dòng, ${droppedQty} PCS`,
    );
    console.log(
      `• ${CFG.skuLoadedSheet}: ${pxkSuperBuildSkuLoadedRows(loadedItems).length} SKU`,
    );
    console.log(
      `• ${CFG.trackingSheet}: đã cập nhật phiếu ${body.ticket?.ticket_no ?? ""}`,
    );
    if (missingNames.length) {
      console.log(
        `• Chưa có Name, đang dùng địa chỉ mặc định: ${missingNames.join(", ")}.`,
      );
    }
  } catch (error) {
    const message = String(error);
    console.log(
      message.includes("Failed to fetch")
        ? `LỖI kết nối API: Failed to fetch — kiểm tra HTTPS, server deploy và CORS. URL: ${CFG.apiUrl}`
        : `LỖI Auto_PXK_Super: ${message}`,
    );
  }
}

const pxkSuperTestables = {
  pxkSuperDimensionHeight,
  pxkSuperPackUnit,
  pxkSuperSafeSheetName,
  pxkSuperTrackingKey,
  pxkSuperValidateUrl,
  pxkSuperLoadedRow,
  pxkSuperCompensationRowIndexes,
  pxkSuperHighlightCompensationRows,
  pxkSuperDroppedRow,
  pxkSuperBuildSkuLoadedRows,
  pxkSuperSkuLoadedRow,
  pxkSuperBalanceWeights,
  pxkSuperJoinContractsFromItems,
};

(
  globalThis as typeof globalThis & {
    __AUTO_PXK_SUPER_TESTABLES__?: typeof pxkSuperTestables;
  }
).__AUTO_PXK_SUPER_TESTABLES__ = pxkSuperTestables;
