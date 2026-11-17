/**
 * ===================================================================
 * Upload_CustomerOrders – Tải đơn đặt hàng từ Excel lên WGR.
 * ===================================================================
 *
 * Dữ liệu bắt đầu từ dòng 2, không phụ thuộc tên tiêu đề ở dòng 1:
 * A Mã đơn hàng (có thể trống) | B SKU | C Mã hợp đồng | D Tên hàng | E Số lượng đơn hàng
 *
 * Các dòng có cùng mã đơn hàng được gom thành một đơn trước khi gửi.
 * Tất cả dòng để trống cột A được gom thành một đơn chưa có mã.
 * Một SKU có thể lặp lại với các mã hợp đồng khác nhau.
 *
 * Với dữ liệu lớn (~100k dòng): đọc Excel một lần (A2:E), gom đơn trong
 * bộ nhớ theo chunk, rồi gửi bulk-import. Tiến độ ghi console ngoài vòng lặp.
 *
 * Cách dùng:
 * 1. Tạo API token tại Dashboard → Hệ thống & Quyền → API Token (Excel).
 * 2. Điền apiToken bên dưới. apiUrl phải là URL HTTPS public (bulk-import).
 * 3. Điền dữ liệu vào sheet cấu hình (mặc định "Đơn đặt hàng") rồi chạy script.
 *    Nếu sheet chưa có / trống / thiếu tiêu đề đúng, script tự tạo sheet mẫu rồi dừng.
 */

const UPLOAD_CUSTOMER_ORDERS_CONFIG = {
  /**
   * Tên sheet nguồn upload.
   * Để trống = dùng sheet đang active (vẫn tạo mẫu theo tenSheetMau nếu bật tuTaoSheetMau).
   */
  tenSheet: "Đơn đặt hàng",
  /**
   * Tự tạo sheet mẫu khi chưa có sheet đúng yêu cầu
   * (không tìm thấy tên cấu hình, sheet trống, hoặc thiếu tiêu đề cột A–E).
   */
  tuTaoSheetMau: true,
  /** Tên sheet mẫu khi tenSheet để trống mà cần tạo mới. */
  tenSheetMau: "Đơn đặt hàng",
  /** Tiêu đề dòng 1 của sheet mẫu. */
  tieuDeMau: [
    "Mã đơn hàng",
    "SKU",
    "Mã hợp đồng",
    "Tên hàng",
    "Số lượng đơn hàng",
  ] as string[],
  /**
   * Dòng ví dụ (từ dòng 2). Xóa/sửa trước khi upload thật.
   * Để [] nếu chỉ muốn tiêu đề trống.
   */
  dongMau: [
    ["WGF25003E1", "901251", "UAS5001-1", "Tủ A", 6],
    ["", "901252", "EUF5002-1", "Tủ B", 4],
  ] as (string | number)[][],
  apiUrl: "https://app.wgr-wh.site/api/v1/warehouse/customer-orders/bulk-import",
  apiToken: "1|hCwtksIBz4Yc7g2wn6aeMO4L6zri5rStWeSYgH9rf7ad4909",
  /** Số dòng xử lý mỗi lần sau khi đã đọc Excel (chỉ dùng trong bộ nhớ). */
  readChunkRows: 5000,
  /** Số đơn tối đa mỗi request bulk. */
  maxOrdersPerRequest: 50,
  /** Tổng dòng SKU/hợp đồng tối đa mỗi request bulk. */
  maxItemsPerRequest: 4000,
  /** Số dòng SKU/hợp đồng tối đa trong một đơn mỗi request. */
  maxItemsPerOrder: 2000,
};

type CustomerOrderExcelValue = string | number | boolean;

type CustomerOrderUploadItem = {
  sku: string;
  contract_no: string;
  product_name: string;
  ordered_quantity: number;
  /** Chỉ dùng trong script để báo dòng trùng; bị bỏ trước khi gửi API. */
  excel_row?: number;
};

type CustomerOrderUploadPayload = {
  order_no: string | null;
  status: "confirmed";
  items: CustomerOrderUploadItem[];
  /** Chunk sau của đơn lớn: bổ sung dòng vào đơn đã tạo ở chunk trước. */
  append_to_existing?: boolean;
};

type CustomerOrderReadResult = {
  orders: CustomerOrderUploadPayload[];
  errors: string[];
  dataRowCount: number;
};

type CustomerOrderBulkApiResponse = {
  message?: string;
  created?: number;
  skipped?: number;
  failed?: number;
  results?: Array<{
    order_no: string | null;
    status: string;
    message: string;
  }>;
};

// @ts-ignore: Tránh cảnh báo trùng tên khi lưu nhiều Office Script trong cùng thư mục.
const customerOrderUpText = (value: CustomerOrderExcelValue): string =>
  String(value ?? "")
    .replace(/[\u00A0\u200B\u200C\u200D\uFEFF]/g, " ")
    .trim();

// @ts-ignore
const customerOrderUpCode = (value: CustomerOrderExcelValue): string =>
  customerOrderUpText(value).replace(/\s+/g, "").toUpperCase();

// @ts-ignore
const customerOrderUpNumber = (value: CustomerOrderExcelValue): number => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  let text = customerOrderUpText(value).replace(/\s/g, "");
  if (!text) return 0;

  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    text =
      lastComma > lastDot
        ? text.replace(/\./g, "").replace(",", ".")
        : text.replace(/,/g, "");
  } else if (lastComma >= 0) {
    text = text.replace(",", ".");
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
};

// @ts-ignore
const customerOrderUpReadRows = (
  rows: CustomerOrderExcelValue[][],
  firstExcelRow = 2,
): CustomerOrderReadResult => {
  const orderMap = new Map<string, CustomerOrderUploadPayload>();
  const errors: string[] = [];
  let dataRowCount = 0;

  rows.forEach((row, index) => {
    const excelRow = firstExcelRow + index;
    const cells = [row[0], row[1], row[2], row[3], row[4]];
    if (cells.every((cell) => customerOrderUpText(cell ?? "") === "")) return;

    dataRowCount += 1;
    const orderNo = customerOrderUpCode(row[0] ?? "");
    const sku = customerOrderUpCode(row[1] ?? "");
    const contractNo = customerOrderUpCode(row[2] ?? "");
    const productName = customerOrderUpText(row[3] ?? "");
    const quantity = customerOrderUpNumber(row[4] ?? "");
    const rowErrors: string[] = [];

    if (!sku) rowErrors.push("thiếu SKU (cột B)");
    if (!contractNo) rowErrors.push("thiếu mã hợp đồng (cột C)");
    if (!productName) rowErrors.push("thiếu tên hàng (cột D)");
    if (!Number.isInteger(quantity) || quantity <= 0) {
      rowErrors.push("số lượng cột E phải là số nguyên dương");
    }

    if (rowErrors.length > 0) {
      errors.push(`Dòng ${excelRow}: ${rowErrors.join("; ")}.`);
      return;
    }

    const orderKey = orderNo || "__WITHOUT_ORDER_NUMBER__";
    let order = orderMap.get(orderKey);
    if (!order) {
      order = { order_no: orderNo || null, status: "confirmed", items: [] };
      orderMap.set(orderKey, order);
    }

    order.items.push({
      sku,
      contract_no: contractNo,
      product_name: productName,
      ordered_quantity: quantity,
      excel_row: excelRow,
    });
  });

  return {
    orders: Array.from(orderMap.values()),
    errors,
    dataRowCount,
  };
};

// @ts-ignore
const customerOrderUpMergeOrders = (
  target: Map<string, CustomerOrderUploadPayload>,
  source: CustomerOrderUploadPayload[],
): void => {
  source.forEach((order) => {
    const orderKey = order.order_no || "__WITHOUT_ORDER_NUMBER__";
    const existing = target.get(orderKey);
    if (!existing) {
      target.set(orderKey, {
        order_no: order.order_no,
        status: "confirmed",
        items: order.items.slice(),
      });
      return;
    }
    existing.items.push(...order.items);
  });
};

// @ts-ignore: Các dòng cùng đơn + SKU + mã hợp đồng được lưu thành dòng riêng (không cộng dồn).
const customerOrderUpDuplicateNotes = (
  orders: CustomerOrderUploadPayload[],
): string[] => {
  const notes: string[] = [];

  orders.forEach((order) => {
    const groups = new Map<string, CustomerOrderUploadItem[]>();
    order.items.forEach((item) => {
      const key = `${item.sku}|${item.contract_no}`;
      const group = groups.get(key) ?? [];
      group.push(item);
      groups.set(key, group);
    });

    groups.forEach((items) => {
      if (items.length < 2) return;
      const lines = items
        .map((item) =>
          item.excel_row
            ? `dòng ${item.excel_row}: ${item.ordered_quantity}`
            : String(item.ordered_quantity),
        )
        .join(", ");
      notes.push(
        `${order.order_no ?? "(không mã đơn)"} – SKU ${items[0].sku} / HĐ ${items[0].contract_no}: ${items.length} dòng (${lines}) → lưu thành ${items.length} dòng riêng.`,
      );
    });
  });

  return notes;
};

// @ts-ignore: Bỏ trường chỉ dùng nội bộ script trước khi gửi API.
const customerOrderUpStripLocalFields = (
  orders: CustomerOrderUploadPayload[],
): CustomerOrderUploadPayload[] =>
  orders.map((order) => ({
    ...order,
    items: order.items.map((item) => ({
      sku: item.sku,
      contract_no: item.contract_no,
      product_name: item.product_name,
      ordered_quantity: item.ordered_quantity,
    })),
  }));

// @ts-ignore: Chia đơn lớn / đơn không mã thành nhiều payload theo giới hạn API.
const customerOrderUpSplitForUpload = (
  orders: CustomerOrderUploadPayload[],
  maxItemsPerOrder: number,
): { uploadOrders: CustomerOrderUploadPayload[]; splitNotes: string[] } => {
  const uploadOrders: CustomerOrderUploadPayload[] = [];
  const splitNotes: string[] = [];

  orders.forEach((order) => {
    if (order.items.length === 0) return;

    if (order.items.length <= maxItemsPerOrder) {
      uploadOrders.push(order);
      return;
    }

    const chunkCount = Math.ceil(order.items.length / maxItemsPerOrder);
    if (order.order_no) {
      splitNotes.push(
        `${order.order_no}: ${order.items.length} dòng → ${chunkCount} request (giới hạn ${maxItemsPerOrder}/đơn).`,
      );
    }

    for (let i = 0; i < order.items.length; i += maxItemsPerOrder) {
      const isFirstChunk = i === 0;
      uploadOrders.push({
        order_no: order.order_no,
        status: "confirmed",
        append_to_existing: order.order_no ? !isFirstChunk : undefined,
        items: order.items.slice(i, i + maxItemsPerOrder),
      });
    }
  });

  return { uploadOrders, splitNotes };
};

// @ts-ignore
const customerOrderUpBuildBatches = (
  orders: CustomerOrderUploadPayload[],
  maxOrdersPerRequest: number,
  maxItemsPerRequest: number,
): CustomerOrderUploadPayload[][] => {
  const batches: CustomerOrderUploadPayload[][] = [];
  let current: CustomerOrderUploadPayload[] = [];
  let currentItems = 0;

  orders.forEach((order) => {
    const itemCount = order.items.length;
    const wouldExceed =
      current.length > 0 &&
      (current.length >= maxOrdersPerRequest ||
        currentItems + itemCount > maxItemsPerRequest);

    // Chunk bổ sung phải đi sau chunk tạo cùng mã — nếu batch hiện tại đã đầy thì đóng batch.
    if (wouldExceed) {
      batches.push(current);
      current = [];
      currentItems = 0;
    }

    current.push(order);
    currentItems += itemCount;
  });

  if (current.length > 0) batches.push(current);
  return batches;
};

// @ts-ignore
const customerOrderUpValidateUrl = (url: string): string | null => {
  const normalized = url.trim().toLowerCase();
  if (!normalized) return "Chưa cấu hình apiUrl.";
  if (!normalized.startsWith("https://")) {
    return "apiUrl phải dùng HTTPS public; Office Scripts không truy cập được localhost hoặc HTTP.";
  }
  if (normalized.includes("localhost") || normalized.includes("127.0.0.1")) {
    return "apiUrl không được dùng localhost hoặc 127.0.0.1.";
  }
  if (!normalized.includes("/customer-orders/bulk-import")) {
    return "apiUrl nên trỏ tới .../customer-orders/bulk-import để upload dữ liệu lớn.";
  }
  return null;
};

const CUSTOMER_ORDER_UP_HEADER_HINTS = [
  ["mã đơn", "ma don", "order"],
  ["sku", "mã hàng", "ma hang"],
  ["hợp đồng", "hop dong", "contract"],
  ["tên hàng", "ten hang", "product", "tên sp"],
  ["số lượng", "so luong", "qty", "quantity"],
];

/** Chuẩn hóa tiêu đề để so khớp gần đúng (bỏ dấu, hạ thường). */
// @ts-ignore
const customerOrderUpNormalizeHeader = (value: CustomerOrderExcelValue): string =>
  customerOrderUpText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d");

/**
 * Sheet đúng yêu cầu khi có tiêu đề A–E khớp gợi ý,
 * hoặc đã có ít nhất một dòng dữ liệu từ dòng 2 (upload vẫn không phụ thuộc tên tiêu đề).
 */
// @ts-ignore
const customerOrderUpSheetHopLe = (
  headerRow: CustomerOrderExcelValue[] | undefined,
  dataRows: CustomerOrderExcelValue[][],
): boolean => {
  const hasData = dataRows.some((row) =>
    [row?.[0], row?.[1], row?.[2], row?.[3], row?.[4]].some(
      (cell) => customerOrderUpText(cell ?? "") !== "",
    ),
  );
  if (hasData) return true;
  if (!headerRow || headerRow.length < 5) return false;

  let matched = 0;
  for (let col = 0; col < 5; col += 1) {
    const header = customerOrderUpNormalizeHeader(headerRow[col] ?? "");
    if (!header) continue;
    if (
      CUSTOMER_ORDER_UP_HEADER_HINTS[col].some((hint) => header.includes(hint))
    ) {
      matched += 1;
    }
  }
  // Chỉ tiêu đề (chưa có dòng dữ liệu) vẫn coi là đúng nếu khớp ≥ 3/5 cột.
  return matched >= 3;
};

/** Tạo/ghi đè sheet mẫu với tiêu đề + dòng ví dụ; trả về tên sheet. */
// @ts-ignore
const customerOrderUpTaoSheetMau = (
  workbook: ExcelScript.Workbook,
  sheetName: string,
  headers: string[],
  sampleRows: (string | number)[][],
): string => {
  const safeName = sheetName.trim().slice(0, 31) || "Đơn đặt hàng";
  let sheet = workbook.getWorksheet(safeName);
  if (!sheet) {
    sheet = workbook.addWorksheet(safeName);
  } else {
    sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    const used = sheet.getUsedRange();
    if (used) used.clear(ExcelScript.ClearApplyTo.all);
  }

  const title =
    headers.length >= 5
      ? headers.slice(0, 5)
      : [
          "Mã đơn hàng",
          "SKU",
          "Mã hợp đồng",
          "Tên hàng",
          "Số lượng đơn hàng",
        ];
  const rows: (string | number)[][] = [
    title,
    ...sampleRows.map((row) => [
      row[0] ?? "",
      row[1] ?? "",
      row[2] ?? "",
      row[3] ?? "",
      row[4] ?? "",
    ]),
  ];
  const lastRow = rows.length;
  sheet.getRange(`A1:E${lastRow}`).setValues(rows);

  const headerRange = sheet.getRange("A1:E1");
  headerRange.getFormat().getFont().setBold(true);
  headerRange
    .getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  headerRange.getFormat().getFill().setColor("#DDEBF7");
  sheet.getFreezePanes().freezeRows(1);
  sheet.getRange("A:E").getFormat().autofitColumns();
  sheet.activate();
  return safeName;
};

// @ts-ignore: Office Scripts cung cấp fetch trên Excel Online.
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

// @ts-ignore
const customerOrderUpParseBulkResponse = (
  bodyText: string,
): CustomerOrderBulkApiResponse => {
  try {
    return JSON.parse(bodyText) as CustomerOrderBulkApiResponse;
  } catch {
    return { message: bodyText };
  }
};

// @ts-ignore
const customerOrderUpPostBulk = async (
  apiUrl: string,
  apiToken: string,
  orders: CustomerOrderUploadPayload[],
): Promise<{
  ok: boolean;
  created: number;
  skipped: number;
  failed: number;
  detail: string;
}> => {
  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${apiToken}`,
      },
      body: JSON.stringify({
        skip_existing: true,
        orders: customerOrderUpStripLocalFields(orders),
      }),
    });
    const bodyText = await response.text();
    const body = customerOrderUpParseBulkResponse(bodyText);
    const created = body.created ?? 0;
    const skipped = body.skipped ?? 0;
    const failed = body.failed ?? 0;

    if (!response.ok) {
      return {
        ok: false,
        created,
        skipped,
        failed: failed > 0 ? failed : orders.length,
        detail: `HTTP ${response.status} – ${body.message ?? bodyText}`,
      };
    }

    return {
      ok: true,
      created,
      skipped,
      failed,
      detail:
        body.message ??
        `created=${created}, skipped=${skipped}, failed=${failed}`,
    };
  } catch (error) {
    return {
      ok: false,
      created: 0,
      skipped: 0,
      failed: orders.length,
      detail: String(error),
    };
  }
};

// @ts-ignore: Hàm entry point của Office Scripts.
async function main(workbook: ExcelScript.Workbook) {
  const config = UPLOAD_CUSTOMER_ORDERS_CONFIG;
  const startedAt = Date.now();
  console.log("Bắt đầu Upload_CustomerOrders...");

  const urlError = customerOrderUpValidateUrl(config.apiUrl);
  if (urlError) {
    console.log(`LỖI cấu hình: ${urlError}`);
    return;
  }
  if (!config.apiToken.trim()) {
    console.log("LỖI: Chưa cấu hình apiToken (Bearer token từ Dashboard).");
    return;
  }

  const configuredName = config.tenSheet.trim();
  const sampleName =
    configuredName || config.tenSheetMau.trim() || "Đơn đặt hàng";
  let sheet = configuredName
    ? workbook.getWorksheet(configuredName)
    : workbook.getActiveWorksheet();

  const ensureSample = (reason: string): never | void => {
    if (!config.tuTaoSheetMau) {
      console.log(`LỖI: ${reason}`);
      return;
    }
    const created = customerOrderUpTaoSheetMau(
      workbook,
      sampleName,
      config.tieuDeMau,
      config.dongMau,
    );
    console.log(
      [
        `Đã tạo sheet mẫu "${created}" vì: ${reason}`,
        "Cột: A Mã đơn hàng | B SKU | C Mã hợp đồng | D Tên hàng | E Số lượng đơn hàng",
        "Điền/sửa dữ liệu từ dòng 2 (có thể xóa dòng ví dụ), rồi chạy lại script.",
      ].join("\n"),
    );
  };

  if (!sheet) {
    ensureSample(
      configuredName
        ? `Không tìm thấy sheet "${configuredName}".`
        : "Không có sheet đang active.",
    );
    return;
  }

  const usedRange = sheet.getUsedRange(true);
  if (!usedRange) {
    ensureSample(`Sheet "${sheet.getName()}" không có dữ liệu.`);
    return;
  }

  const lastExcelRow = usedRange.getRowIndex() + usedRange.getRowCount();
  if (lastExcelRow < 1) {
    ensureSample(`Sheet "${sheet.getName()}" không có dữ liệu.`);
    return;
  }

  // Sheet lớn: chỉ đọc dòng 1 để kiểm tra mẫu; dữ liệu thật đọc một lần bên dưới.
  if (lastExcelRow < 2) {
    const headerOnly = sheet.getRange("A1:E1").getValues()[0] as
      | CustomerOrderExcelValue[]
      | undefined;
    if (!customerOrderUpSheetHopLe(headerOnly, [])) {
      ensureSample(
        `Sheet "${sheet.getName()}" chưa đúng yêu cầu (thiếu tiêu đề A–E).`,
      );
      return;
    }
    console.log(
      `Sheet "${sheet.getName()}" đã có tiêu đề mẫu nhưng chưa có dữ liệu từ dòng 2. Hãy điền rồi chạy lại.`,
    );
    return;
  }

  const totalDataRows = lastExcelRow - 1;
  const readChunkRows = Math.max(500, config.readChunkRows || 5000);
  console.log(
    `Sheet "${sheet.getName()}": ${totalDataRows} dòng (A2:E${lastExcelRow}). Đọc một lần, xử lý chunk ${readChunkRows} dòng trong bộ nhớ...`,
  );

  // Một lần getValues/getTexts — tránh read API trong vòng lặp (Office Scripts performance).
  const dataRange = sheet.getRange(`A2:E${lastExcelRow}`);
  const values = dataRange.getValues() as CustomerOrderExcelValue[][];
  const texts = dataRange.getTexts() as string[][];
  const allRows = values.map((row, index) =>
    row.map((value, column) => {
      if (column === 4 && typeof value === "number") return value;
      return texts[index]?.[column] ?? value;
    }),
  );

  // Sheet có vùng dùng từ dòng 2 nhưng toàn dòng trống / không khớp yêu cầu cột.
  if (!customerOrderUpSheetHopLe(undefined, allRows)) {
    ensureSample(
      `Sheet "${sheet.getName()}" chưa đúng yêu cầu (không có dòng dữ liệu hợp lệ từ dòng 2).`,
    );
    return;
  }

  const orderMap = new Map<string, CustomerOrderUploadPayload>();
  const errors: string[] = [];
  let dataRowCount = 0;

  for (let offset = 0; offset < allRows.length; offset += readChunkRows) {
    const chunkRows = allRows.slice(offset, offset + readChunkRows);
    const chunkResult = customerOrderUpReadRows(chunkRows, 2 + offset);
    dataRowCount += chunkResult.dataRowCount;
    errors.push(...chunkResult.errors);
    customerOrderUpMergeOrders(orderMap, chunkResult.orders);
  }

  console.log(
    `Đã đọc xong: dòng có dữ liệu=${dataRowCount} | đơn tạm=${orderMap.size} | lỗi=${errors.length}`,
  );

  if (errors.length > 0) {
    const errorPreview = errors.slice(0, 50);
    const more =
      errors.length > 50 ? [`... và ${errors.length - 50} lỗi khác.`] : [];
    console.log(
      [
        `DỪNG UPLOAD: phát hiện ${errors.length} dòng không hợp lệ. Hãy sửa dữ liệu rồi chạy lại.`,
        ...errorPreview,
        ...more,
      ].join("\n"),
    );
    return;
  }

  const groupedOrders = Array.from(orderMap.values());
  if (groupedOrders.length === 0) {
    console.log(
      `Không tìm thấy dữ liệu hợp lệ từ dòng 2 trên sheet "${sheet.getName()}".`,
    );
    return;
  }

  const duplicateNotes = customerOrderUpDuplicateNotes(groupedOrders);
  if (duplicateNotes.length > 0) {
    console.log(
      [
        `Ghi chú: ${duplicateNotes.length} cặp SKU + mã hợp đồng lặp lại trong cùng đơn (không cộng dồn):`,
        ...duplicateNotes.slice(0, 50),
        ...(duplicateNotes.length > 50
          ? [`... và ${duplicateNotes.length - 50} cặp khác.`]
          : []),
      ].join("\n"),
    );
  }

  const maxItemsPerOrder = Math.max(1, config.maxItemsPerOrder || 2000);
  const { uploadOrders, splitNotes } = customerOrderUpSplitForUpload(
    groupedOrders,
    maxItemsPerOrder,
  );

  if (splitNotes.length > 0) {
    console.log(
      [
        `Đã tách ${splitNotes.length} đơn lớn để upload hết (giữ nguyên mã đơn):`,
        ...splitNotes.slice(0, 20),
        ...(splitNotes.length > 20
          ? [`... và ${splitNotes.length - 20} đơn khác.`]
          : []),
        "Lưu ý: nếu đơn lớn đã tạo một phần rồi chạy lại, hãy xóa đơn đó trên hệ thống trước để tránh trùng dòng bổ sung.",
      ].join("\n"),
    );
  }

  if (uploadOrders.length === 0) {
    console.log("Không còn đơn hợp lệ để upload sau khi chia payload.");
    return;
  }

  const batches = customerOrderUpBuildBatches(
    uploadOrders,
    Math.max(1, config.maxOrdersPerRequest || 50),
    Math.max(1, config.maxItemsPerRequest || 4000),
  );

  console.log(
    `Đã gom ${dataRowCount} dòng → ${groupedOrders.length} đơn gốc → ${uploadOrders.length} payload. Gửi ${batches.length} request bulk...`,
  );

  const apiToken = config.apiToken.trim();
  let created = 0;
  let skipped = 0;
  let failed = 0;
  const failedDetails: string[] = [];
  const batchLogs: string[] = [];

  for (let i = 0; i < batches.length; i += 1) {
    const batch = batches[i];
    const batchItems = batch.reduce(
      (sum, order) => sum + order.items.length,
      0,
    );

    const outcome = await customerOrderUpPostBulk(
      config.apiUrl.trim(),
      apiToken,
      batch,
    );
    created += outcome.created;
    skipped += outcome.skipped;
    failed += outcome.failed;

    batchLogs.push(
      `Batch ${i + 1}/${batches.length}: ${batch.length} đơn, ${batchItems} dòng — ${outcome.detail} | tích lũy created=${created}, skipped=${skipped}, failed=${failed}`,
    );

    if (!outcome.ok) {
      failedDetails.push(`Batch ${i + 1}: ${outcome.detail}`);
    }
  }

  const elapsedSec = Math.round((Date.now() - startedAt) / 1000);
  console.log(
    [
      ...batchLogs,
      `HOÀN TẤT trong ${elapsedSec}s: created=${created}, skipped=${skipped}, failed=${failed}.`,
      ...(failedDetails.length > 0
        ? ["Chi tiết lỗi:", ...failedDetails.slice(0, 20)]
        : []),
      skipped > 0
        ? "Đơn đã tồn tại được bỏ qua (skip_existing=true) — chạy lại không nhân đôi đơn có mã."
        : "",
    ]
      .filter((line) => line !== "")
      .join("\n"),
  );
}

const customerOrderUpTestables = {
  customerOrderUpText,
  customerOrderUpCode,
  customerOrderUpNumber,
  customerOrderUpReadRows,
  customerOrderUpMergeOrders,
  customerOrderUpDuplicateNotes,
  customerOrderUpStripLocalFields,
  customerOrderUpSplitForUpload,
  customerOrderUpBuildBatches,
  customerOrderUpValidateUrl,
  customerOrderUpNormalizeHeader,
  customerOrderUpSheetHopLe,
  customerOrderUpTaoSheetMau,
};

(
  globalThis as typeof globalThis & {
    __UPLOAD_CUSTOMER_ORDERS_TESTABLES__?: typeof customerOrderUpTestables;
  }
).__UPLOAD_CUSTOMER_ORDERS_TESTABLES__ = customerOrderUpTestables;
