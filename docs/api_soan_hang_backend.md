# API soạn hàng từ backend (tương thích Auto_Mitsuha)

Các endpoint đều cần đăng nhập Sanctum hoặc Bearer token.

## Sinh danh sách đề xuất

`GET /api/warehouse/picking-wave-planner`

Office Script có thể gửi cấu hình riêng cho một lần chạy mà không thay đổi cấu hình đã lưu:

`POST /api/warehouse/picking-wave-planner/preview`

```json
{
  "configuration": {
    "inventory_filter_enabled": true,
    "default_min_inventory_threshold": 5,
    "market_min_inventory_thresholds": { "EU": 8 },
    "market_pair_batch_sizes": { "UE": 10 },
    "market_codes": ["EU"],
    "receipt_date_filter": { "days": [], "months": [], "years": [] },
    "included_skus": [],
    "excluded_skus": [],
    "max_skus": 0,
    "require_ab_pair": true,
    "include_warning_items": true,
    "prioritize_completed_contracts": true,
    "prioritize_classification": false,
    "default_classification_order": ["S", "T", "M", "C", "F"],
    "classification_order_by_market": {}
  }
}
```

Response:

```json
{
  "format": "auto_mitsuha_backend_v1",
  "data": [
    {
      "stt": 1,
      "product_id": 12,
      "contract_id": 34,
      "location_id": 5,
      "contract_no": "EUF2026-001",
      "sku": "912345",
      "pair_key": "912345|EUF2026-001",
      "product_name": "Tên hàng",
      "dimensions": "40*30*20",
      "order_quantity": 100,
      "previously_shipped_quantity": 20,
      "requested_quantity": 40,
      "unit_weight": 2.5,
      "total_weight": 100,
      "unit_volume": 0.024,
      "total_volume": 0.96,
      "location_code": "A1",
      "current_stock": 45,
      "available_stock": 40,
      "allocated_stock": 5,
      "dropped_quantity": 8,
      "dropped_label": "RỚT 8/40",
      "dropped_dispatch_no": "PXK-20260903-001",
      "dropped_dispatch_date": "2026-09-03",
      "dropped_dispatch_reference": "PXK-20260903-001 — 03/09/2026",
      "pending_production": -10,
      "eligible": true,
      "warning_reason": null
    }
  ],
  "summary": {
    "total_lines": 1,
    "eligible_lines": 1,
    "warning_lines": 0,
    "total_quantity": 40,
    "total_weight": 100,
    "total_volume": 0.96
  },
  "configuration": {}
}
```

Các trường nghiệp vụ tương ứng với dòng output Auto_Mitsuha: hợp đồng, SKU, tên hàng, kích thước, số lượng đơn, đã xuất/đã giữ, số lượng cần soạn, trọng lượng, thể tích, vị trí, mã phân loại, tồn thực tế, chưa hoàn thành và ngày nhập mới nhất. Các ID backend được trả thêm để có thể gửi thẳng dòng đạt chuẩn sang `POST /api/warehouse/picking-waves`.

Hàng rớt được tính từ các Pre-packing-list ở trạng thái `converted` đã liên kết với PXK chưa hủy: số lượng đã soạn trên PPL trừ số lượng thực xuất theo SKU. Dòng có hàng rớt được ưu tiên lên đầu và trả thêm `dropped_quantity`, `dropped_label`, `dropped_dispatch_no`, `dropped_dispatch_date`, `dropped_dispatch_reference`, `dropped_invoice_no`. Auto Mitsuha Super ghi nhãn vào cột R, chuỗi `Số PXK — ngày xuất kho` vào cột S và invoice của PXK đó vào cột T.

Mỗi dòng còn có `order_no`: mã đơn hàng khách (`customer_orders.order_no`) gắn với cặp SKU + hợp đồng qua phân bổ đơn hàng, nhiều mã nối bằng khoảng trắng, rỗng nếu chưa gắn đơn. Auto Mitsuha Super ghi vào cột U; các cột Mã phân loại → Ngày nhập mới nhất dời sang V → Z.

`pair_key` là khóa đối chiếu nghiệp vụ chuẩn, luôn ghép theo thứ tự **SKU + mã hợp đồng**: `SKU|CONTRACT_NO` (ví dụ `912345|EUF2026-001`). SKU và mã hợp đồng được viết hoa, bỏ khoảng trắng trước khi ghép. Số phiếu pre-packing/PXK không nằm trong khóa này; nó chỉ là dữ liệu lịch sử hiển thị khi truy vết.

Quy ước áp dụng thống nhất trong hệ thống:

- API, Excel và import/export: `SKU|CONTRACT_NO`.
- Cache, tra cứu nội bộ và composite index: `product_id|contract_id`.
- Khi có vị trí kho: `product_id|contract_id|location_id`.
- Mã pre-packing, Picking Wave và PXK là mã chứng từ, không tham gia khóa nhận diện SKU–hợp đồng.

`eligible=false` là dòng cảnh báo màu vàng tương đương vùng chưa đạt của Auto_Mitsuha và không nên tự động đưa vào wave.

## Đọc và lưu cấu hình

- `GET /api/warehouse/picking-wave-planner/configuration`
- `PUT /api/warehouse/picking-wave-planner/configuration`

Quyền đọc/sinh đề xuất: `warehouse.picking-waves.view`.

Quyền sửa cấu hình: `warehouse.picking-waves.manage-config`.

Trang giao diện: `/dashboard/warehouse/picking-waves/configuration`.
