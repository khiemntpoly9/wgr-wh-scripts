# wgr-work

## Sync_PrePackingList

Script `finished-goods/Sync_PrePackingList.ts` kéo danh sách tủ tải trước hiện tại trên hệ thống về sheet
`pre-packing-list` (format Auto_PXK, cột A→R), để file Excel khớp với hệ thống sau khi bổ sung, sửa hoặc xoá SKU trên
web.

### Cấu hình

Sửa `SYNC_PRE_PACKING_CONFIG` ở đầu script sau khi dán vào Excel Online:

| Khoá | Ý nghĩa |
|------|---------|
| `apiUrl` | `https://<domain>/api/v1/warehouse/pre-packing-lists/sheet-data`. Phải là HTTPS public. |
| `apiToken` | Token tạo tại Dashboard → API Token (Excel). Tài khoản cần quyền xem pre-packing list. |
| `listNoCell` | Ô chứa số phiếu, mặc định `D4` (`NO: 000631`). |
| `listNo` | Ghi đè số phiếu thủ công; để trống thì đọc từ `listNoCell`. |
| `markAddedColor`, `markUpdatedColor` | Màu đánh dấu dòng bổ sung và dòng thay đổi. |
| `lotSheetName` | Sheet tra cứu số lô (mặc định `lot-no`): cột A hợp đồng, cột B số lô, dòng 1 là tiêu đề. |
| `missingLotColor` | Màu ô K của dòng bổ sung không tìm được số lô. |
| `unitColumns` | ĐVT cho dòng bổ sung: G = `qtyUnitDefault`, J = `packUnitPackage` khi chiều cao bằng `packageHeight`, ngược lại `packUnitCarton`. |
| `summaryLimit` | Số dòng tối đa mỗi danh sách trong tóm tắt. |

Không lưu token thật vào file trên repo; để `apiToken: ""`.

### Kết quả trên sheet

- Dòng có trên hệ thống mà Excel chưa có được chèn ngay dưới dòng cuối cùng cùng hợp đồng, hoặc cuối vùng dữ liệu nếu
  là hợp đồng mới.
- Dòng bổ sung điền ĐVT và số lô như `Auto_PXK.ts`: G = `PCS`; J = `PACKAGE` khi chiều cao (số thứ 3 cột E) là 1,1,
  ngược lại `箱/ THÙNG CARTON`. Số lô trống được tra trong sheet `lot-no` theo hợp đồng, sau đó lấy số lô 10 chữ số của
  dòng cùng hợp đồng gốc (EUF6001-1, EUF6001 3 → EUF6001). Không tìm được thì ô K tô đỏ và liệt kê trong tóm tắt.
- Dòng đã có được ghi đè các cột hệ thống quản lý: F, H, I, K (khi hệ thống có số lô), L, M, N, O, P, Q (khi hệ thống có
  ghi chú), R (khi PMC đã duyệt). Cột B, C, D, E, G, J giữ nguyên.
- Dòng không còn trên hệ thống bị xoá. Dòng thiếu SKU hoặc hợp đồng (dòng trống, dòng ghi chú) không bị chạm tới.
- STT được đánh lại, công thức SUM ở dòng tổng được viết lại theo vùng dữ liệu mới.
- Tóm tắt trong khung output liệt kê dòng đã thêm, đã xoá và từng ô bị ghi đè (giá trị cũ → mới).

### Chú thích màu

| Màu | Ý nghĩa |
|-----|---------|
| Xanh lá nhạt `#C6EFCE` | Dòng bổ sung từ hệ thống |
| Vàng nhạt `#FFEB9C` | Dòng có ít nhất một ô bị ghi đè theo hệ thống |

Màu chỉ phản ánh lần đồng bộ gần nhất. Lần chạy sau gỡ các dòng đang mang đúng hai màu này về không nền rồi mới tô lại.
Màu do người dùng tự tô trên dòng không đổi được giữ nguyên.

### Đồng bộ trước khi upload lại

`Upload_PrePackingList.ts` từ chối upload (HTTP 409) khi phiếu đã có SKU bổ sung trên hệ thống mà file Excel thiếu, và
liệt kê các SKU đó. Cách xử lý:

1. Chạy `Sync_PrePackingList` để kéo các SKU bổ sung về Excel, rồi upload lại.
2. Hoặc đặt `replaceAll: true` trong `UPLOAD_PRE_PACKING_CONFIG` nếu muốn thay thế toàn bộ phiếu theo file. Các SKU bổ
   sung bị loại bỏ và được liệt kê trong kết quả upload.

### Token cũ

Bản `Upload_PrePackingList.ts` trên repo trước đây có token hard-code. Hãy thu hồi token đó tại Dashboard → API Token
và tạo token mới để dán vào script khi dùng.

## Nhập kho một bước (Upload_Nami_Super)

Script `finished-goods/Upload_Nami_Super.ts` đọc thẳng sheet "Thành phẩm", lấy các lần nhập kho (Nhập kho 1/2/3) của
kỳ đã chọn và gửi lên `POST /api/v1/warehouse/receipts/nami-direct-import`. Không cần chạy `Auto_Nami` trước và không
cần sheet `KQ_NhapKho_*`.

### Cấu hình

Sửa `UPLOAD_NAMI_SUPER_CONFIG` ở đầu script:

| Khoá | Ý nghĩa |
|------|---------|
| `apiToken` | Token tạo tại Dashboard → API Token (Excel). Tài khoản cần quyền nhập kho. Không lưu token thật vào repo. |
| `ngay` | Ngày chọn `dd/mm/yyyy`; để trống = hôm nay theo giờ Việt Nam. |
| `cheDo` | `"day"` theo ngày, `"week"` thứ Hai → Chủ nhật chứa ngày chọn, `"month"` cả tháng. |
| `traThongSoSheet` | `false` thì không đọc sheet "2026". |
| `kichThuocDot` | Số lần nhập tối đa mỗi request (mặc định 300). |

### Thông số sản phẩm

Thông số đã có trên hệ thống (trọng lượng, kích thước, thể tích) được giữ nguyên. Sheet "2026" chỉ dùng để điền phần
hệ thống còn thiếu, tra theo mã → tên → tên gần đúng như `Auto_Nami`. Nhật ký liệt kê SKU lấy thông số từ sheet (nhất là
SKU theo tên gần đúng, cần kiểm tra lại) và SKU vẫn chưa có thông số.

### Đánh dấu lỗi

Lần nhập bị bỏ qua (SL trống/0/chữ, thiếu SKU hoặc mã hợp đồng, tổng K + M + O nhỏ hơn SL) hoặc bị hệ thống từ chối
(không khớp đơn hàng) được tô màu `#F4B084` ở đúng cặp ô ngày + SL: J:K (lần 1), L:M (lần 2), N:O (lần 3). Lý do in trong
nhật ký.

Danh sách ô đã tô và màu cũ được lưu trong sheet ẩn `_UploadNamiSuper`. Lần chạy sau trả các ô đó về màu cũ trước khi tô
lại; ô bạn đã tô sang màu khác không bị đụng tới. Không xoá hoặc sửa sheet ẩn này.

### Chạy lại

Chạy lại cùng kỳ là an toàn: lần nhập đã có trên hệ thống được tính "đã có từ trước" và không bị nhập trùng, kể cả khi
đã tải bằng `Upload_Nami`. Khi một đợt bị lỗi giữa chừng, các đợt trước đã được ghi; sửa lỗi rồi chạy lại cả kỳ. Gặp
HTTP 409 nghĩa là đang có lần nhập kho khác chạy, đợi 1–2 phút rồi thử lại.

### Báo cáo

Muốn xem báo cáo dạng `KQ_NhapKho_*` của kỳ vừa tải, chạy `Auto_Nami_Supper` với cùng ngày/chế độ.

`Auto_Nami` + `Upload_Nami` là cách cũ, vẫn dùng được.
