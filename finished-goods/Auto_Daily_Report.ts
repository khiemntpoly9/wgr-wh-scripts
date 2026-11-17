/**
 * ===================================================================
 * Auto_Daily_Report – Báo cáo nhập / xuất hàng hằng ngày.
 * ===================================================================
 *
 * Cách dùng mỗi ngày:
 * 1. Copy sheet báo cáo hôm trước thành sheet mới, đặt tên DDMM (vd "2909").
 * 2. Sửa A1 sang ngày cần báo cáo (vd 29/09/2026) — để trống thì lấy theo tên sheet.
 * 3. Mở sheet đó và chạy script. Tên sheet và A1 lệch ngày thì script dừng.
 *
 * Kết quả trên sheet đang mở:
 * - Xuất (phiếu xuất kho), mỗi cont một số hạng:
 *   - T9:  số lượng xuất trong ngày, vd =1307+1483 (=0 nếu không xuất).
 *   - N9:  thể tích xuất trong ngày, vd =198.64+331.82.
 *   - T11: lũy kế số cont tháng, nối số cont của ngày vào cuối, vd =3+5+...+2.
 * - Nhập (phiếu nhập kho đã hoàn tất), tổng của ngày:
 *   - H9:  số thùng nhập trong ngày, vd =1250.
 *   - B9:  thể tích nhập trong ngày, vd =85.32.
 * - Lũy kế tháng T10 / N10 / H10 / B10: công thức cũ + TỔNG của ô ngày
 *   (T9 / N9 / H9 / B9) hôm trước, vd =T9+...+2790, =N9+...+530.46.
 * - Sang tháng mới: T10 = =T9, N10 = =N9, H10 = =H9, B10 = =B9, T11 bắt đầu lại.
 * - Tồn kho quy cont (thể tích / 68) theo mã ở hàng 3 (B3 EUF, C3 EUS, ... U3 UET):
 *   - B4:U4 tồn đầu ngày = giá trị B7:U7 của hôm trước (chỉ copy 1 lần mỗi ngày).
 *   - B5:U5 thể tích xuất / 68 theo mã, vd =198.64/68 (mã không xuất thì để trống).
 *   - B6:U6 thể tích nhập / 68 theo mã (mã không nhập thì để trống).
 *   - B7:U7 tồn cuối = =B4-B5+B6; V4 và V7 = tổng hàng (=SUM(B7:U7)), V5/V6 không đụng.
 * - Dấu thập phân: script ghi công thức dạng chuẩn (dấu chấm), Excel tự hiển thị
 *   dấu phẩy nếu máy cài vùng Việt Nam (vd =N9+198,64+331,82).
 *
 * Ô trạng thái (mặc định AZ1, chữ trắng) lưu JSON để chạy lại cùng ngày
 * không bị cộng trùng. Ô này được copy theo sheet — không xóa.
 *
 * Chèn thêm cột (vd mã EUC, PLC): địa chỉ ở trên là mặc định. Đặt Name phạm vi
 * sheet cho các ô (xem config, vd rpt_out_qty_day = T9, rpt_stock_codes = B3:U3,
 * rpt_state = AZ1) thì chèn / xóa cột không làm script ghi lệch ô. Cột mã mới
 * phải chèn vào giữa vùng rpt_stock_codes (trước cột tổng) để Name tự giãn.
 *
 * Token cần quyền xem phiếu xuất kho và phiếu nhập kho.
 * Office Scripts chạy trên Microsoft cloud: URL API phải là HTTPS public.
 */

const AUTO_DAILY_REPORT_CONFIG = {
  dispatchApiUrl:
    "https://app.wgr-wh.site/api/v1/warehouse/dispatch-orders/report",
  receiptApiUrl: "https://app.wgr-wh.site/api/v1/warehouse/receipts/report",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909", // Bearer token Sanctum từ Dashboard → API Token (Excel)
  /*
   * Mỗi ô / vùng dưới đây tìm theo Name trước (Formulas → Name Manager, Scope = sheet
   * để Name đi theo khi copy sheet); chưa có Name thì dùng địa chỉ fallback.
   * Excel tự dời Name khi chèn / xóa hàng, cột nên script không bị lệch ô.
   */
  dateCell: { name: "rpt_date", fallback: "A1" },
  /** Ô ngày / ô lũy kế tháng cho từng chỉ số. */
  cells: {
    /** Số lượng xuất (PCS), mỗi cont một số hạng. */
    outQty: {
      /** Trong ngày, vd =1307+1483. */
      daily: { name: "rpt_out_qty_day", fallback: "T9" },
      /** Lũy kế tháng = công thức cũ + tổng ô ngày hôm trước. */
      month: { name: "rpt_out_qty_month", fallback: "T10" },
    },
    /** Thể tích xuất (m³), mỗi cont một số hạng. */
    outVolume: {
      /** Trong ngày, vd =198.64+331.82. */
      daily: { name: "rpt_out_vol_day", fallback: "N9" },
      /** Lũy kế tháng. */
      month: { name: "rpt_out_vol_month", fallback: "N10" },
    },
    /** Số thùng nhập (phiếu nhập đã hoàn tất), tổng của ngày. */
    inCartons: {
      /** Trong ngày, vd =1250. */
      daily: { name: "rpt_in_ctn_day", fallback: "H9" },
      /** Lũy kế tháng. */
      month: { name: "rpt_in_ctn_month", fallback: "H10" },
    },
    /** Thể tích nhập (m³), tổng của ngày. */
    inVolume: {
      /** Trong ngày, vd =85.32. */
      daily: { name: "rpt_in_vol_day", fallback: "B9" },
      /** Lũy kế tháng. */
      month: { name: "rpt_in_vol_month", fallback: "B10" },
    },
  },
  /** Lũy kế số cont xuất trong tháng. */
  outContCell: { name: "rpt_out_cont", fallback: "T11" },
  /** Số chữ số thập phân khi ghi thể tích (m³). */
  volumeDecimals: 2,
  /**
   * Bảng tồn kho quy ra cont (thể tích / containerCbm) theo mã thị trường + phân loại.
   * codes là cả hàng mã (EUF, EUS, ... UET); chèn cột mã mới vào GIỮA vùng này thì
   * Excel tự giãn Name, script tự nhận mã mới. total là ô cột tổng trên cùng hàng mã.
   * Các hàng tồn đầu / xuất / nhập / tồn cuối tính lệch so với hàng mã.
   */
  stock: {
    /**
     * Cả hàng mã (EUF, EUS, ... UET). Xác định cột đầu / cuối và hàng gốc của bảng;
     * cột có mã trống thì giữ nguyên nội dung cũ.
     */
    codes: { name: "rpt_stock_codes", fallback: "B3:U3" },
    /**
     * Ô tiêu đề cột tổng (sau mã cuối). Chỉ lấy cột để ghi =SUM vào hàng tồn đầu
     * và tồn cuối; nội dung ô không được đọc.
     */
    total: { name: "rpt_stock_total", fallback: "V3" },
    openOffset: 1,
    outOffset: 2,
    inOffset: 3,
    closeOffset: 4,
    containerCbm: 68,
  },
  stateCell: { name: "rpt_state", fallback: "AZ1" },
  /**
   * Chỉ dùng cho lần chạy đầu tiên (chưa có ô trạng thái): ngày của số liệu
   * đang nằm ở các ô ngày (T9...), định dạng dd/mm/yyyy. Để trống = coi là cùng tháng với A1.
   */
  initialPrevDate: "",
  /**
   * Tên sheet dạng DDMM (vd "2909" = 29/09, "0110" = 01/10; "2909 (2)" cũng được).
   * true: tên sheet và A1 lệch ngày thì dừng; A1 trống thì lấy ngày từ tên sheet.
   * Tên sheet không theo dạng DDMM thì bỏ qua bước kiểm tra.
   */
  checkSheetName: true,
};

type DailyReportMetricKey = "outQty" | "outVolume" | "inCartons" | "inVolume";

const DAILY_REPORT_METRIC_LABELS: { [key in DailyReportMetricKey]: string } = {
  outQty: "Số lượng xuất",
  outVolume: "Thể tích xuất",
  inCartons: "Số thùng nhập",
  inVolume: "Thể tích nhập",
};

type DailyReportRow = {
  id?: number;
  dispatch_no?: string | null;
  dispatch_date?: string | null;
  container_no?: string | null;
  total_quantity?: number | string | null;
  total_volume?: number | string | null;
};

type DailyReportByCode = {
  code?: string | null;
  total_volume?: number | string | null;
};

type DailyReportResponse = {
  message?: string;
  data?: {
    selected_date?: string | null;
    rows?: DailyReportRow[];
    by_code?: DailyReportByCode[];
  };
};

type DailyReportReceiptMarket = {
  market_code?: string | null;
  classifications?: {
    classification_code?: string | null;
    total_volume?: number | string | null;
  }[];
};

type DailyReportReceiptResponse = {
  message?: string;
  data?: {
    date_from?: string | null;
    date_to?: string | null;
    summary?: {
      receipt_count?: number | null;
      carton_quantity?: number | string | null;
      total_volume?: number | string | null;
    };
    by_market?: DailyReportReceiptMarket[];
  };
};

type DailyReportStockInput = {
  /** Mã ở hàng codeRow theo thứ tự cột (ô trống = ""). */
  codes: string[];
  /** Cột tương ứng với từng mã (vd "B"). */
  columns: string[];
  openRow: number;
  outRow: number;
  inRow: number;
  outByCode: { [code: string]: number };
  inByCode: { [code: string]: number };
  containerCbm: number;
  volumeDecimals: number;
};

type DailyReportStockResult = {
  outFormulas: string[];
  inFormulas: string[];
  closeFormulas: string[];
  /** Mã có thể tích nhưng không có cột ở hàng mã. */
  unmatchedOut: string[];
  unmatchedIn: string[];
};

type DailyReportState = {
  /** Ngày (yyyy-mm-dd) của số liệu đang ghi ở các ô ngày. */
  date: string;
  /** Ngày của số liệu ô ngày trước khi script ghi đè, null nếu không rõ. */
  prevDate: string | null;
  /** Giá trị ô ngày trước khi script ghi đè (số của ngày prevDate), theo chỉ số. */
  prevDaily: { [key: string]: number };
  /** Công thức ô lũy kế tháng trước khi script sửa, theo chỉ số. */
  baseMonth: { [key: string]: string };
  /** Công thức T11 trước khi script sửa; null = không có trong trạng thái. */
  baseCount: string | null;
  /** Đã copy tồn cuối ngày (hàng 7) lên tồn đầu ngày (hàng 4) cho ngày date. */
  stockOpened?: boolean;
};

type DailyReportContainers = {
  /** Số lượng từng cont (bỏ cont số lượng 0). */
  quantities: number[];
  /** Thể tích từng cont đã làm tròn (bỏ cont thể tích 0). */
  volumes: number[];
};

type DailyReportMetricInput = {
  key: DailyReportMetricKey;
  dailyCell: string;
  decimals: number;
  terms: number[];
  currentDailyValue: unknown;
  currentMonthFormula: unknown;
};

type DailyReportComputeInput = {
  reportDate: string;
  state: DailyReportState | null;
  initialPrevDate: string | null;
  metrics: DailyReportMetricInput[];
  currentCountFormula: unknown;
  containerCount: number;
};

type DailyReportMetricResult = {
  daily: string;
  month: string;
  appendedPrev: number;
};

type DailyReportComputeResult = {
  metrics: { [key: string]: DailyReportMetricResult };
  count: string;
  state: DailyReportState;
  monthReset: boolean;
};

const dailyReportPad2 = (value: number): string =>
  value < 10 ? `0${value}` : String(value);

const dailyReportIso = (
  year: number,
  month: number,
  day: number,
): string | null => {
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return `${year}-${dailyReportPad2(month)}-${dailyReportPad2(day)}`;
};

/** Excel serial, Date hoặc text dd/mm/yyyy (hoặc yyyy-mm-dd) → yyyy-mm-dd. */
const dailyReportParseDate = (value: unknown): string | null => {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return dailyReportIso(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate(),
    );
  }
  if (typeof value === "number") {
    if (!isFinite(value) || value < 1) return null;
    const date = new Date(
      Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000,
    );
    return dailyReportIso(
      date.getUTCFullYear(),
      date.getUTCMonth() + 1,
      date.getUTCDate(),
    );
  }

  const text = String(value ?? "").trim();
  if (!text) return null;
  const iso = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    return dailyReportIso(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  }
  const dmy = text.match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
  if (dmy) {
    return dailyReportIso(Number(dmy[3]), Number(dmy[2]), Number(dmy[1]));
  }
  if (/^\d+(\.\d+)?$/.test(text)) {
    return dailyReportParseDate(Number(text));
  }
  return null;
};

type DailyReportSheetDay = { day: number; month: number };

/** Tên sheet "2909" → { day: 29, month: 9 }; không đúng dạng DDMM thì null. */
const dailyReportParseSheetName = (
  name: string,
): DailyReportSheetDay | null => {
  const match = String(name ?? "").match(/^\s*(\d{2})(\d{2})(?!\d)/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { day, month };
};

const dailyReportSheetNameMatches = (
  iso: string,
  sheetDay: DailyReportSheetDay,
): boolean =>
  Number(iso.slice(5, 7)) === sheetDay.month &&
  Number(iso.slice(8, 10)) === sheetDay.day;

/**
 * Suy ra năm cho ngày lấy từ tên sheet dựa trên ngày tham chiếu (ngày hôm trước):
 * tháng trên tên sheet nhỏ hơn tháng tham chiếu thì coi là năm sau (12 → 01).
 */
const dailyReportDateFromSheetName = (
  sheetDay: DailyReportSheetDay,
  referenceIso: string,
): string | null => {
  let year = Number(referenceIso.slice(0, 4));
  if (sheetDay.month < Number(referenceIso.slice(5, 7))) year += 1;
  return dailyReportIso(year, sheetDay.month, sheetDay.day);
};

const dailyReportIsoToSerial = (iso: string): number => {
  const [year, month, day] = iso.split("-").map((part) => Number(part));
  return Math.round(
    (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000,
  );
};

const dailyReportFormatVn = (iso: string): string => {
  const [year, month, day] = iso.split("-");
  return `${day}/${month}/${year}`;
};

const dailyReportIsSameMonth = (a: string, b: string): boolean =>
  a.slice(0, 7) === b.slice(0, 7);

/** Số hoặc text số; text chỉ có dấu phẩy (vd "198,64") coi phẩy là thập phân. */
const dailyReportNumber = (value: unknown): number => {
  if (typeof value === "number") return isFinite(value) ? value : 0;
  let text = String(value ?? "")
    .trim()
    .replace(/\s/g, "");
  if (text.includes(",") && !text.includes(".")) text = text.replace(/,/g, ".");
  const num = Number(text);
  return isFinite(num) ? num : 0;
};

const dailyReportRound = (value: number, decimals: number): number => {
  const factor = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * factor) / factor;
};

/**
 * Gom số lượng và thể tích theo cont (theo thứ tự xuất hiện).
 * Cont có số lượng 0 bị bỏ; thể tích làm tròn volumeDecimals, bỏ cont thể tích 0.
 */
const dailyReportGroupContainers = (
  rows: DailyReportRow[],
  volumeDecimals: number,
): DailyReportContainers => {
  const order: string[] = [];
  const totals: { [key: string]: { qty: number; volume: number } } = {};
  rows.forEach((row, index) => {
    const container = String(row.container_no ?? "")
      .trim()
      .toUpperCase();
    const key = container
      ? `C:${container}`
      : `D:${String(row.dispatch_no ?? "").trim() || row.id || index}`;
    if (!(key in totals)) {
      order.push(key);
      totals[key] = { qty: 0, volume: 0 };
    }
    totals[key].qty += dailyReportNumber(row.total_quantity);
    totals[key].volume += dailyReportNumber(row.total_volume);
  });
  const kept = order.map((key) => totals[key]).filter((item) => item.qty > 0);
  return {
    quantities: kept.map((item) => item.qty),
    volumes: kept
      .map((item) => dailyReportRound(item.volume, volumeDecimals))
      .filter((volume) => volume > 0),
  };
};

/** Tổng nhập của ngày → số hạng cho H9 / B9 (bỏ nếu bằng 0). */
const dailyReportReceiptTerms = (
  summary: { carton_quantity?: unknown; total_volume?: unknown } | undefined,
  volumeDecimals: number,
): { cartons: number[]; volumes: number[] } => {
  const cartons = dailyReportNumber(summary?.carton_quantity);
  const volume = dailyReportRound(
    dailyReportNumber(summary?.total_volume),
    volumeDecimals,
  );
  return {
    cartons: cartons > 0 ? [cartons] : [],
    volumes: volume > 0 ? [volume] : [],
  };
};

const dailyReportBuildSumFormula = (terms: number[]): string =>
  terms.length ? `=${terms.join("+")}` : "=0";

const dailyReportFormulaText = (formula: unknown): string => {
  if (formula === null || formula === undefined) return "";
  return String(formula).trim();
};

/** Nối thêm một số hạng vào cuối công thức cộng; ô trống hoặc =0 thì bắt đầu mới. */
const dailyReportAppendTerm = (formula: unknown, term: number): string => {
  const text = dailyReportFormulaText(formula);
  if (text === "" || text === "=" || text === "=0" || text === "0") {
    return `=${term}`;
  }
  return text.startsWith("=") ? `${text}+${term}` : `=${text}+${term}`;
};

/** Trạng thái bản đầu (chỉ có xuất) lưu field phẳng theo ô T9/T10/T11/N9/N10. */
const DAILY_REPORT_LEGACY_STATE_KEYS: {
  prev: { [field: string]: DailyReportMetricKey };
  base: { [field: string]: DailyReportMetricKey };
} = {
  prev: { prevT9Value: "outQty", prevN9Value: "outVolume" },
  base: { baseT10: "outQty", baseN10: "outVolume" },
};

const dailyReportParseState = (value: unknown): DailyReportState | null => {
  const text = String(value ?? "").trim();
  if (!text.startsWith("{")) return null;
  try {
    const raw = JSON.parse(text) as { [field: string]: unknown };
    const date = dailyReportParseDate(raw.date);
    if (!date) return null;

    const prevDaily: { [key: string]: number } = {};
    const baseMonth: { [key: string]: string } = {};
    const rawPrev = raw.prevDaily;
    const rawBase = raw.baseMonth;
    if (rawPrev && typeof rawPrev === "object") {
      Object.keys(rawPrev).forEach((key) => {
        prevDaily[key] = dailyReportNumber(
          (rawPrev as { [k: string]: unknown })[key],
        );
      });
    }
    if (rawBase && typeof rawBase === "object") {
      Object.keys(rawBase).forEach((key) => {
        baseMonth[key] = dailyReportFormulaText(
          (rawBase as { [k: string]: unknown })[key],
        );
      });
    }
    Object.keys(DAILY_REPORT_LEGACY_STATE_KEYS.prev).forEach((field) => {
      const key = DAILY_REPORT_LEGACY_STATE_KEYS.prev[field];
      if (
        raw[field] !== undefined &&
        raw[field] !== null &&
        !(key in prevDaily)
      ) {
        prevDaily[key] = dailyReportNumber(raw[field]);
      }
    });
    Object.keys(DAILY_REPORT_LEGACY_STATE_KEYS.base).forEach((field) => {
      const key = DAILY_REPORT_LEGACY_STATE_KEYS.base[field];
      if (
        raw[field] !== undefined &&
        raw[field] !== null &&
        !(key in baseMonth)
      ) {
        baseMonth[key] = dailyReportFormulaText(raw[field]);
      }
    });

    const rawCount = raw.baseCount ?? raw.baseT11;
    return {
      date,
      prevDate: raw.prevDate ? dailyReportParseDate(raw.prevDate) : null,
      prevDaily,
      baseMonth,
      baseCount:
        rawCount === undefined || rawCount === null
          ? null
          : dailyReportFormulaText(rawCount),
      ...(raw.stockOpened === true ? { stockOpened: true } : {}),
    };
  } catch {
    return null;
  }
};

const dailyReportSerializeState = (state: DailyReportState): string =>
  JSON.stringify(state);

/**
 * Công thức lũy kế tháng: gốc + tổng của ô ngày hôm trước (chỉ số tổng,
 * không chép chi tiết). Sang tháng mới hoặc gốc chưa có công thức thì = ô ngày.
 */
const dailyReportMonthFormula = (
  base: string,
  dailyCell: string,
  prevValue: number,
  monthReset: boolean,
): string => {
  const dailyRef = `=${dailyCell}`;
  if (monthReset) return dailyRef;
  const start = base.startsWith("=") ? base : dailyRef;
  return prevValue > 0 ? dailyReportAppendTerm(start, prevValue) : start;
};

/**
 * Tính công thức mới cho các ô ngày / lũy kế tháng / T11 và trạng thái cần lưu,
 * không đụng workbook. Chạy lại cùng ngày dùng gốc trong trạng thái; chỉ số
 * chưa có trong trạng thái (vd vừa thêm vào script) thì lấy giá trị đang có trên sheet.
 */
const dailyReportCompute = (
  input: DailyReportComputeInput,
): DailyReportComputeResult => {
  const state = input.state;
  const rerun = state !== null && state.date === input.reportDate;
  const prevDate = rerun
    ? state!.prevDate
    : (state?.date ?? input.initialPrevDate);
  const monthReset =
    prevDate !== null && !dailyReportIsSameMonth(prevDate, input.reportDate);

  const prevDaily: { [key: string]: number } = {};
  const baseMonth: { [key: string]: string } = {};
  const metrics: { [key: string]: DailyReportMetricResult } = {};

  input.metrics.forEach((metric) => {
    const base =
      rerun && metric.key in state!.baseMonth
        ? state!.baseMonth[metric.key]
        : dailyReportFormulaText(metric.currentMonthFormula);
    const prevValue = dailyReportRound(
      rerun && metric.key in state!.prevDaily
        ? state!.prevDaily[metric.key]
        : dailyReportNumber(metric.currentDailyValue),
      metric.decimals,
    );
    prevDaily[metric.key] = prevValue;
    baseMonth[metric.key] = base;
    metrics[metric.key] = {
      daily: dailyReportBuildSumFormula(metric.terms),
      month: dailyReportMonthFormula(
        base,
        metric.dailyCell,
        prevValue,
        monthReset,
      ),
      appendedPrev: !monthReset && prevValue > 0 ? prevValue : 0,
    };
  });

  const baseCount =
    rerun && state!.baseCount !== null
      ? state!.baseCount
      : dailyReportFormulaText(input.currentCountFormula);
  const countStart = monthReset ? "" : baseCount;
  const count = input.containerCount
    ? dailyReportAppendTerm(countStart, input.containerCount)
    : countStart || "=0";

  return {
    metrics,
    count,
    state: {
      date: input.reportDate,
      prevDate,
      prevDaily,
      baseMonth,
      baseCount,
    },
    monthReset,
  };
};

const dailyReportColumnIndex = (column: string): number =>
  column
    .trim()
    .toUpperCase()
    .split("")
    .reduce((index, char) => index * 26 + (char.charCodeAt(0) - 64), 0);

const dailyReportColumnName = (index: number): string => {
  let name = "";
  let rest = index;
  while (rest > 0) {
    const mod = (rest - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    rest = Math.floor((rest - 1) / 26);
  }
  return name;
};

/** "B", "U" → ["B", "C", ..., "U"]. */
const dailyReportColumnRange = (first: string, last: string): string[] => {
  const columns: string[] = [];
  for (
    let index = dailyReportColumnIndex(first);
    index <= dailyReportColumnIndex(last);
    index++
  ) {
    columns.push(dailyReportColumnName(index));
  }
  return columns;
};

const dailyReportCode = (value: unknown): string =>
  String(value ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");

/** by_code của API phiếu xuất → { EUF: 12.5, ... }. */
const dailyReportDispatchVolumeByCode = (
  rows: DailyReportByCode[] | undefined,
): { [code: string]: number } => {
  const result: { [code: string]: number } = {};
  (rows ?? []).forEach((row) => {
    const code = dailyReportCode(row.code);
    if (!code) return;
    result[code] = (result[code] ?? 0) + dailyReportNumber(row.total_volume);
  });
  return result;
};

/** by_market[].classifications[] của API phiếu nhập → { EUF: 12.5, ... }. */
const dailyReportReceiptVolumeByCode = (
  markets: DailyReportReceiptMarket[] | undefined,
): { [code: string]: number } => {
  const result: { [code: string]: number } = {};
  (markets ?? []).forEach((market) => {
    const marketCode = dailyReportCode(market.market_code);
    (market.classifications ?? []).forEach((item) => {
      const code = marketCode + dailyReportCode(item.classification_code);
      if (!code) return;
      result[code] = (result[code] ?? 0) + dailyReportNumber(item.total_volume);
    });
  });
  return result;
};

/**
 * Công thức hàng xuất / nhập (thể tích / cbm cont) và hàng tồn cuối (=đầu - xuất + nhập)
 * cho từng cột mã. Mã không có thể tích thì ô xuất / nhập để trống.
 */
const dailyReportStockRows = (
  input: DailyReportStockInput,
): DailyReportStockResult => {
  const known: { [code: string]: true } = {};
  const toFormula = (volume: number | undefined): string => {
    const rounded = dailyReportRound(volume ?? 0, input.volumeDecimals);
    return rounded > 0 ? `=${rounded}/${input.containerCbm}` : "";
  };
  const outFormulas: string[] = [];
  const inFormulas: string[] = [];
  const closeFormulas: string[] = [];

  input.codes.forEach((rawCode, index) => {
    const code = dailyReportCode(rawCode);
    const column = input.columns[index];
    if (!code) {
      outFormulas.push("");
      inFormulas.push("");
      closeFormulas.push("");
      return;
    }
    known[code] = true;
    outFormulas.push(toFormula(input.outByCode[code]));
    inFormulas.push(toFormula(input.inByCode[code]));
    closeFormulas.push(
      `=${column}${input.openRow}-${column}${input.outRow}+${column}${input.inRow}`,
    );
  });

  const unmatched = (byCode: { [code: string]: number }) =>
    Object.keys(byCode)
      .filter(
        (code) =>
          !known[code] &&
          dailyReportRound(byCode[code], input.volumeDecimals) > 0,
      )
      .sort();

  return {
    outFormulas,
    inFormulas,
    closeFormulas,
    unmatchedOut: unmatched(input.outByCode),
    unmatchedIn: unmatched(input.inByCode),
  };
};

type DailyReportRef = { name: string; fallback: string };

/** "'Sheet 1'!$T$9" → "T9". */
const dailyReportLocalAddress = (address: string): string => {
  const text = String(address ?? "");
  const bang = text.lastIndexOf("!");
  return (bang >= 0 ? text.slice(bang + 1) : text).replace(/\$/g, "");
};

/**
 * Name phạm vi sheet → Name phạm vi workbook (chỉ khi trỏ vào chính sheet này)
 * → địa chỉ fallback. Ghi tên Name không tìm thấy vào missing.
 */
const dailyReportResolveRange = (
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  ref: DailyReportRef,
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

const dailyReportValidateUrl = (url: string): string | null => {
  const trimmed = url.trim();
  if (!trimmed) return "Chưa cấu hình URL API.";
  if (!/^https:\/\//i.test(trimmed)) {
    return `URL API phải là HTTPS public (Office Scripts chạy trên cloud): ${trimmed}`;
  }
  return null;
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

/** GET JSON có Bearer token; lỗi thì trả message để main log và dừng. */
async function dailyReportFetchJson<T extends { message?: string }>(
  url: string,
  token: string,
  label: string,
): Promise<{ body: T | null; error: string | null }> {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${token}`,
    },
  });
  const bodyText = await response.text();
  let body: T;
  try {
    body = JSON.parse(bodyText) as T;
  } catch {
    return {
      body: null,
      error: `LỖI ${label}: API trả dữ liệu không phải JSON (${response.status}).`,
    };
  }
  if (!response.ok) {
    return {
      body: null,
      error: `LỖI ${label} ${response.status}: ${body.message ?? bodyText}`,
    };
  }
  return { body, error: null };
}

// @ts-ignore: Mỗi Office Script chạy độc lập nên có thể dùng tên main.
async function main(workbook: ExcelScript.Workbook) {
  const CFG = AUTO_DAILY_REPORT_CONFIG;
  const urlError =
    dailyReportValidateUrl(CFG.dispatchApiUrl) ??
    dailyReportValidateUrl(CFG.receiptApiUrl);
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

  let initialPrevDate: string | null = null;
  if (CFG.initialPrevDate.trim()) {
    initialPrevDate = dailyReportParseDate(CFG.initialPrevDate);
    if (!initialPrevDate) {
      console.log(
        `LỖI cấu hình: initialPrevDate "${CFG.initialPrevDate}" phải có dạng dd/mm/yyyy.`,
      );
      return;
    }
  }

  const sheet = workbook.getActiveWorksheet();
  const sheetName = sheet.getName();
  const missingNames: string[] = [];
  const resolve = (ref: DailyReportRef) =>
    dailyReportResolveRange(workbook, sheet, ref, missingNames);
  const addressOf = (range: ExcelScript.Range) =>
    dailyReportLocalAddress(range.getAddress());

  const stateRange = resolve(CFG.stateCell);
  const state = dailyReportParseState(stateRange.getValue());

  const dateRange = resolve(CFG.dateCell);
  const dateAddress = addressOf(dateRange);
  const dateCellText = String(dateRange.getText() ?? "").trim();
  let reportDate =
    dailyReportParseDate(dateRange.getValue()) ??
    dailyReportParseDate(dateCellText);
  const sheetDay = CFG.checkSheetName
    ? dailyReportParseSheetName(sheetName)
    : null;

  if (!reportDate && sheetDay && !dateCellText) {
    const referenceIso =
      state?.date ??
      initialPrevDate ??
      dailyReportIso(
        new Date().getFullYear(),
        new Date().getMonth() + 1,
        new Date().getDate(),
      )!;
    reportDate = dailyReportDateFromSheetName(sheetDay, referenceIso);
    if (reportDate) {
      dateRange.setValue(dailyReportIsoToSerial(reportDate));
      dateRange.setNumberFormat("dd/mm/yyyy");
      console.log(
        `• ${dateAddress} trống: đã điền ${dailyReportFormatVn(reportDate)} theo tên sheet "${sheetName}".`,
      );
    }
  }

  if (!reportDate) {
    console.log(
      `LỖI: Ô ${dateAddress} trên sheet "${sheetName}" không phải ngày hợp lệ (vd 28/09/2026).`,
    );
    return;
  }

  if (sheetDay && !dailyReportSheetNameMatches(reportDate, sheetDay)) {
    console.log(
      `LỖI: Tên sheet "${sheetName}" là ngày ${dailyReportPad2(sheetDay.day)}/${dailyReportPad2(sheetDay.month)}, ` +
        `nhưng ${dateAddress} là ${dailyReportFormatVn(reportDate)}. ` +
        `Hãy sửa ${dateAddress} hoặc tên sheet cho khớp rồi chạy lại.`,
    );
    return;
  }

  if (state && state.date > reportDate) {
    console.log(
      `LỖI: Sheet này đã có số liệu ngày ${dailyReportFormatVn(state.date)}, ` +
        `mới hơn ngày ${dailyReportFormatVn(reportDate)} ở ${dateAddress}. ` +
        "Hãy copy đúng sheet của ngày trước đó.",
    );
    return;
  }

  const date = encodeURIComponent(reportDate);
  const dispatchUrl =
    CFG.dispatchApiUrl.replace(/\?.*$/, "") + `?date=${date}&period=day`;
  const receiptUrl =
    CFG.receiptApiUrl.replace(/\?.*$/, "") +
    `?date_from=${date}&date_to=${date}&group_by=day`;
  console.log(
    `Đang lấy phiếu xuất kho và phiếu nhập kho ngày ${dailyReportFormatVn(reportDate)}...`,
  );

  try {
    const dispatch = await dailyReportFetchJson<DailyReportResponse>(
      dispatchUrl,
      CFG.apiToken,
      "API phiếu xuất",
    );
    if (dispatch.error || !dispatch.body) {
      console.log(dispatch.error);
      return;
    }
    const selectedDate = dispatch.body.data?.selected_date;
    if (selectedDate && selectedDate !== reportDate) {
      console.log(
        `LỖI: API phiếu xuất trả ngày ${selectedDate}, không khớp ngày ${reportDate}.`,
      );
      return;
    }

    const receipt = await dailyReportFetchJson<DailyReportReceiptResponse>(
      receiptUrl,
      CFG.apiToken,
      "API phiếu nhập",
    );
    if (receipt.error || !receipt.body) {
      console.log(receipt.error);
      return;
    }

    const out = dailyReportGroupContainers(
      dispatch.body.data?.rows ?? [],
      CFG.volumeDecimals,
    );
    const inbound = dailyReportReceiptTerms(
      receipt.body.data?.summary,
      CFG.volumeDecimals,
    );
    const terms: { [key in DailyReportMetricKey]: number[] } = {
      outQty: out.quantities,
      outVolume: out.volumes,
      inCartons: inbound.cartons,
      inVolume: inbound.volumes,
    };
    const decimals: { [key in DailyReportMetricKey]: number } = {
      outQty: 0,
      outVolume: CFG.volumeDecimals,
      inCartons: 0,
      inVolume: CFG.volumeDecimals,
    };
    const keys = Object.keys(CFG.cells) as DailyReportMetricKey[];
    const dailyRanges = keys.map((key) => resolve(CFG.cells[key].daily));
    const monthRanges = keys.map((key) => resolve(CFG.cells[key].month));
    const metricAddresses = keys.map((_, index) => ({
      daily: addressOf(dailyRanges[index]),
      month: addressOf(monthRanges[index]),
    }));
    const countRange = resolve(CFG.outContCell);

    const result = dailyReportCompute({
      reportDate,
      state,
      initialPrevDate,
      metrics: keys.map((key, index) => ({
        key,
        dailyCell: metricAddresses[index].daily,
        decimals: decimals[key],
        terms: terms[key],
        currentDailyValue: dailyRanges[index].getValue(),
        currentMonthFormula: monthRanges[index].getFormula(),
      })),
      currentCountFormula: countRange.getFormula(),
      containerCount: out.quantities.length,
    });

    // setFormula dùng cú pháp invariant (dấu chấm thập phân); Excel tự hiển thị
    // dấu phẩy trên máy cài vùng dùng dấu phẩy.
    keys.forEach((key, index) => {
      dailyRanges[index].setFormula(result.metrics[key].daily);
      monthRanges[index].setFormula(result.metrics[key].month);
    });
    countRange.setFormula(result.count);

    const stock = CFG.stock;
    const codesRange = resolve(stock.codes);
    const firstColIndex = codesRange.getColumnIndex() + 1;
    const firstCol = dailyReportColumnName(firstColIndex);
    const lastCol = dailyReportColumnName(
      firstColIndex + codesRange.getColumnCount() - 1,
    );
    const totalCol = dailyReportColumnName(
      resolve(stock.total).getColumnIndex() + 1,
    );
    const codeRow = codesRange.getRowIndex() + 1;
    const openRow = codeRow + stock.openOffset;
    const outRow = codeRow + stock.outOffset;
    const inRow = codeRow + stock.inOffset;
    const closeRow = codeRow + stock.closeOffset;
    const stockColumns = dailyReportColumnRange(firstCol, lastCol);
    const rowAddress = (row: number) => `${firstCol}${row}:${lastCol}${row}`;
    const stockCodes = sheet
      .getRange(rowAddress(codeRow))
      .getTexts()[0]
      .map((text) => dailyReportCode(text));
    const stockRows = dailyReportStockRows({
      codes: stockCodes,
      columns: stockColumns,
      openRow,
      outRow,
      inRow,
      outByCode: dailyReportDispatchVolumeByCode(dispatch.body.data?.by_code),
      inByCode: dailyReportReceiptVolumeByCode(receipt.body.data?.by_market),
      containerCbm: stock.containerCbm,
      volumeDecimals: CFG.volumeDecimals,
    });

    // Tồn cuối hôm trước → tồn đầu hôm nay, chỉ một lần mỗi ngày (chạy lại thì giữ nguyên).
    const rerun = state !== null && state.date === reportDate;
    const stockAlreadyOpened = rerun && state!.stockOpened === true;
    if (!stockAlreadyOpened) {
      const closeValues = sheet.getRange(rowAddress(closeRow)).getValues();
      const openRange = sheet.getRange(rowAddress(openRow));
      const openValues = openRange.getValues();
      openRange.setValues([
        closeValues[0].map((value, index) =>
          stockCodes[index] ? value : openValues[0][index],
        ),
      ]);
    }

    // Cột không có mã ở hàng mã thì giữ nguyên nội dung cũ.
    const writeStockRow = (row: number, formulas: string[]) => {
      const range = sheet.getRange(rowAddress(row));
      const existing = range.getFormulas()[0];
      range.setFormulas([
        formulas.map((formula, index) =>
          stockCodes[index] ? formula : String(existing[index] ?? ""),
        ),
      ]);
    };
    writeStockRow(outRow, stockRows.outFormulas);
    writeStockRow(inRow, stockRows.inFormulas);
    writeStockRow(closeRow, stockRows.closeFormulas);
    [openRow, closeRow].forEach((row) => {
      sheet
        .getRange(`${totalCol}${row}`)
        .setFormula(`=SUM(${rowAddress(row)})`);
    });

    stateRange.setValue(
      dailyReportSerializeState({ ...result.state, stockOpened: true }),
    );
    stateRange.getFormat().getFont().setColor("#FFFFFF");

    const sum = (values: number[], places: number) =>
      dailyReportRound(
        values.reduce((total, value) => total + value, 0),
        places,
      );
    const monthNote = result.monthReset ? " (tháng mới)" : "";
    console.log("=========================================");
    console.log("  HOÀN TẤT AUTO_DAILY_REPORT");
    console.log("=========================================");
    console.log(
      `• Xuất ${dailyReportFormatVn(reportDate)}: ${out.quantities.length} cont, ` +
        `${sum(out.quantities, 0)} PCS, ${sum(out.volumes, CFG.volumeDecimals)} m³`,
    );
    console.log(
      `• Nhập ${dailyReportFormatVn(reportDate)}: ${receipt.body.data?.summary?.receipt_count ?? 0} phiếu, ` +
        `${sum(inbound.cartons, 0)} thùng, ${sum(inbound.volumes, CFG.volumeDecimals)} m³`,
    );
    keys.forEach((key, index) => {
      console.log(
        `• ${DAILY_REPORT_METRIC_LABELS[key]}: ${metricAddresses[index].daily} ${result.metrics[key].daily} | ` +
          `${metricAddresses[index].month} ${result.metrics[key].month}${monthNote}`,
      );
    });
    console.log(
      `• Số cont xuất tháng: ${addressOf(countRange)} ${result.count}`,
    );
    console.log(
      `• Tồn kho quy cont (/${stock.containerCbm}, cột ${firstCol}:${lastCol}, tổng ${totalCol}): ` +
        (stockAlreadyOpened
          ? `giữ nguyên tồn đầu ngày hàng ${openRow} (chạy lại cùng ngày)`
          : `đã copy tồn cuối hàng ${closeRow} lên tồn đầu hàng ${openRow}`) +
        `; cập nhật xuất hàng ${outRow}, nhập hàng ${inRow}, tồn cuối hàng ${closeRow}.`,
    );
    if (missingNames.length) {
      console.log(
        `• Chưa có Name, đang dùng địa chỉ mặc định: ${missingNames.join(", ")}.`,
      );
    }
    if (dispatch.body.data?.by_code === undefined && out.quantities.length) {
      console.log(
        `• CẢNH BÁO: API phiếu xuất chưa trả by_code (backend chưa cập nhật) — hàng ${outRow} đang ghi 0.`,
      );
    }
    if (stockRows.unmatchedOut.length || stockRows.unmatchedIn.length) {
      console.log(
        `• CẢNH BÁO: mã không có cột ở hàng ${codeRow}, thể tích chưa được ghi — ` +
          `xuất: ${stockRows.unmatchedOut.join(", ") || "không"}; ` +
          `nhập: ${stockRows.unmatchedIn.join(", ") || "không"}.`,
      );
    }
    if (state && state.date === reportDate) {
      console.log("• Chạy lại cùng ngày: đã tính lại, không cộng trùng.");
    }
  } catch (error) {
    const message = String(error);
    console.log(
      message.includes("Failed to fetch")
        ? "LỖI kết nối API: Failed to fetch — kiểm tra HTTPS, server deploy và CORS."
        : `LỖI Auto_Daily_Report: ${message}`,
    );
  }
}

const dailyReportTestables = {
  dailyReportParseDate,
  dailyReportParseSheetName,
  dailyReportSheetNameMatches,
  dailyReportDateFromSheetName,
  dailyReportIsoToSerial,
  dailyReportIsSameMonth,
  dailyReportNumber,
  dailyReportGroupContainers,
  dailyReportReceiptTerms,
  dailyReportColumnRange,
  dailyReportDispatchVolumeByCode,
  dailyReportReceiptVolumeByCode,
  dailyReportStockRows,
  dailyReportBuildSumFormula,
  dailyReportAppendTerm,
  dailyReportParseState,
  dailyReportSerializeState,
  dailyReportCompute,
  dailyReportValidateUrl,
  dailyReportLocalAddress,
};

(
  globalThis as typeof globalThis & {
    __AUTO_DAILY_REPORT_TESTABLES__?: typeof dailyReportTestables;
  }
).__AUTO_DAILY_REPORT_TESTABLES__ = dailyReportTestables;
