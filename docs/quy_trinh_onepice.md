# Quy trình OnePice chuẩn hóa tính số lượng soạn/xuất

## 1) Mục tiêu và phạm vi

Tài liệu này định nghĩa **một logic duy nhất** để:

- Tính số lượng hàng **có thể soạn** (gợi ý cho danh sách soạn).
- Tính số lượng **có thể xuất tối đa** cho một dòng xuất cụ thể.
- Kiểm tra một dòng nhập xuất có **đủ điều kiện xuất** hay không.
- Tách rõ 2 kết luận nghiệp vụ:
  - `can_ship_now`: có thể xuất ngay ở lần hiện tại hay không.
  - `can_wait_for_next_ship`: nếu chưa xuất lần này thì trạng thái có hợp lệ để chờ lần sau hay không.

Quy trình này không phụ thuộc chữ cái cột (`H`, `AJ`, `AC`...) vì mỗi script có thể map cột khác nhau.  
Mỗi script chỉ cần map về cùng một bộ trường dữ liệu chuẩn.

---

## 1.1 Phân vai 3 script (bắt buộc)

Để tránh nhầm luồng, OnePice chốt phân vai như sau:

- `Auto_Mitsuha.ts` (**soạn hàng**):
  - Mục tiêu: tạo danh sách soạn tối ưu theo rule tồn kho/ngưỡng/chống kẹt.
  - Đầu ra: `duocSoan`, `soLuongSoanGoiY`, và dữ liệu phục vụ gom nhóm/xuất theo cặp.
  - Được phép tự tối ưu số lượng gợi ý theo rule.

- `Auto_Doraemon.ts` (**kiểm tra trước xuất**):
  - Mục tiêu: kiểm tra phiếu trước khi chạy xuất thật.
  - Đầu ra: `OK/Lỗi/Cảnh báo`, có thể kèm số gợi ý điều chỉnh.
  - Bắt buộc kiểm tra quy tắc SKU hậu tố `A/B`: không cho phép xuất tách rời 1 vế.
  - Không ghi xuất kho vào sheet đích.

- `Auto_Taki.ts` (**thực thi xuất kho**):
  - Mục tiêu: ghi xuất trực tiếp vào sheet theo đúng số lượng trên phiếu.
  - Vai trò: validator + executor, không tự tối ưu lại số lượng như luồng soạn.
  - Nếu số lượng trên phiếu không đạt điều kiện thì từ chối ghi và trả lý do.

Nguyên tắc phân vai:

- `Mitsuha` quyết định **nên soạn bao nhiêu**.
- `Doraemon` quyết định **phiếu có hợp lệ trước khi xuất không**.
- `Taki` quyết định **có ghi xuất đúng số trên phiếu được hay không**.

## 1.2 Chính sách SKU bổ sung (đã chốt)

Mức áp dụng: **A (mềm)**.

- SKU nằm trong `locSKU` được phép **hiển thị để người vận hành thấy**.
- SKU bổ sung **không được vượt qua các rule cứng**:
  - Không vi phạm ngưỡng tồn kho/chống kẹt kho.
  - Không vi phạm xuất theo cặp (nếu thị trường có áp).
  - Không vi phạm dữ liệu nền (`tongNhap > slDonHang`, `daXuat > slDonHang`, ...).
- Tóm lại: SKU bổ sung chỉ là cơ chế ưu tiên hiển thị, **không phải cơ chế nới điều kiện xuất/soạn**.

---

## 2) Chuẩn dữ liệu đầu vào (schema logic)

Mỗi dòng sản phẩm cần có tối thiểu các trường sau:

- `maHang`: mã hàng.
- `hopDong`: mã hợp đồng/thị trường.
- `soLuongDonHang` (SLDH): số lượng đơn hàng.
- `tongNhap`: tổng nhập kho đã hoàn thành.
- `tonKhoVatLy`: tồn kho vật lý hiện hành.
- `dsXuatDaGhi`: danh sách các lần xuất đã ghi (có thể nhiều slot).
- `soLuongDeXuat` (tùy ngữ cảnh): lượng người dùng muốn xuất ở dòng hiện tại.

Các tham số cấu hình dùng chung (theo Mitsuha mới):

- `tonKhoToiThieu` (ngưỡng mặc định, ví dụ: `5`).
- `tonKhoTheoThiTruong` (ghi đè theo 2 ký tự đầu HĐ, ví dụ: `EU -> 8`).
- `locThiTruong` (lọc thị trường; để trống thì lấy tất cả).
- `xuatTheoCapTheoThiTruong` (rule gom cặp theo thị trường, áp ở bước hậu xử lý).
- `locSKU` (danh sách SKU bổ sung; có ngoại lệ riêng khi qua rule ngưỡng).
- Số lượng cho phép số thập phân kỹ thuật và làm tròn `4` chữ số để tránh sai số JS.

---

## 3) Hàm chuẩn hóa số liệu

## 3.1 Hàm số an toàn

Với mọi ô số:

- Rỗng -> `0`.
- Không parse được số -> lỗi dữ liệu (chặn script) hoặc quy ước `0` tùy tầng xử lý.
- Sau mọi phép cộng/trừ quan trọng -> làm tròn `toFixed(4)` để tránh sai số JS.

## 3.2 Hàm ngưỡng tồn kho theo hợp đồng (Mitsuha)

```
nguongTonKhoToiThieu = CONFIG.tonKhoTheoThiTruong[prefixHopDong] ?? CONFIG.tonKhoToiThieu
```

Trong đó `prefixHopDong` là 2 ký tự đầu của hợp đồng đã làm sạch.

## 3.3 Hàm tính xuất đã ghi

```
daXuatThucTe = sum(dsXuatDaGhiDaLamSach)
```

Không phụ thuộc số lượng cột xuất (3 cột, 4 cột, ...).

## 3.4 Hàm tính chưa hoàn thành chuẩn

```
chtThucTe = max(0, soLuongDonHang - tongNhap)
chtBang = giaTriCotChuaHoanThanh; neu khong hop le hoac am => chtThucTe
chuaHoanThanh = max(chtBang, chtThucTe)
```

Ghi chú: lấy số lớn hơn để không đánh giá thiếu phần đang dở dang (tránh xuất quá tay).

## 3.5 Hàm tính còn lại của đơn chuẩn

```
conLaiCuaDon = soLuongDonHang - daXuatThucTe
```

---

## 4) Tính số lượng có thể soạn (gợi ý)

Áp dụng khi tạo danh sách soạn tự động:

## Bước S1: Dữ liệu nền

- `tonKhoVatLy` luôn tự tính lại: `tongNhap - tongXuat` (không tin tuyệt đối vào cột tồn kho hiển thị).
- `daXuatThucTe = tongXuat`.
- `conLaiCuaDon = soLuongDonHang - daXuatThucTe`.
- Nếu `conLaiCuaDon <= 0` thì bỏ qua, trừ `SKU bổ sung`.
- Nếu `conLaiCuaDon < 0` thì ép về `0` để an toàn tính toán.

## Bước S2: Trần theo tồn và đơn

```
soLuongSoanToiDa = min(tonKhoVatLy, conLaiCuaDon)
```

## Bước S3: Rule ngưỡng tồn kho

Nếu bật `locTonKho` và `tonKhoVatLy < nguongTonKhoToiThieu`:

- Nếu `chuaHoanThanh > 0` và **không phải SKU bổ sung** -> bỏ qua dòng.
- Nếu là SKU bổ sung thì được đi tiếp (ngoại lệ vận hành của Mitsuha).

## Bước S4: Chống kẹt kho cuối đơn

Nếu `0 < chuaHoanThanh < nguongTonKhoToiThieu`:

```
tranChongKet = tonKhoVatLy + chuaHoanThanh - nguongTonKhoToiThieu
soLuongSoanToiDa = min(soLuongSoanToiDa, tranChongKet)
```

## Bước S5: Xử lý dưới ngưỡng sau chống kẹt

- Nếu `soLuongSoanToiDa <= 0` -> không soạn.
- Nếu `soLuongSoanToiDa < nguongTonKhoToiThieu` và **không phải SKU bổ sung** -> không soạn.

Kết quả:

```
soLuongSoanGoiY = round(soLuongSoanToiDa, 4)
```

---

## 5) Hậu xử lý theo nhóm (xuất theo cặp)

Sau khi có danh sách tạm đủ điều kiện soạn:

- Gom theo nhóm `Hợp đồng + SKU`.
- Tính tổng `soLuongSoanGoiY` của nhóm.
- Nếu thị trường có `xuatTheoCapTheoThiTruong > 0` và tổng nhóm chưa đủ ngưỡng cặp thì loại nhóm đó.
- Với SKU dạng `6 số + A/B` (ví dụ `901019A`, `901019B`):
  - Gom theo nhóm `Hợp đồng + SKU gốc (6 số)`.
  - Chỉ cho qua khi nhóm có đủ cả `A` và `B`.
  - Nếu thiếu một vế thì chặn toàn bộ nhóm để tránh tách cặp.

Quy tắc này áp dụng nhất quán:

- `Mitsuha`: chặn khỏi danh sách đủ điều kiện và đẩy xuống vùng cảnh báo.
- `Doraemon`: báo lỗi phiếu nếu dữ liệu xuất có `A` mà thiếu `B` (hoặc ngược lại) trong cùng nhóm.

---

## 6) Kiểm tra có thể soạn (full validation checklist)

Với từng dòng nguồn:

1. Parse dữ liệu số hợp lệ.
2. Tính chuẩn: `daXuatThucTe`, `chuaHoanThanh`, `conLaiCuaDon`, `nguongTonKhoToiThieu`.
3. Tính `soLuongSoanToiDa = min(tonKhoVatLy, conLaiCuaDon)`.
4. Áp rule ngưỡng + chống kẹt + ngoại lệ SKU bổ sung.
5. Nếu pass -> đưa vào danh sách tạm, chờ hậu xử lý theo cặp.

## 6.1 Kết luận đầu ra cho luồng soạn

Với mỗi dòng:

- `duocSoan`: có/không.
- `soLuongSoanGoiY`: số lượng gợi ý sau toàn bộ rule.
- `lyDo`: lý do bỏ qua (nếu có).

---

## 7) Quy tắc đồng nhất giữa các script

Mọi script (`Auto_Mitsuha`, `Auto_Taki`, `Auto_Conan`, ...) phải cùng:

- Dùng cùng công thức `chuaHoanThanh = max(0, SLDH - TongNhap)`.
- Dùng cùng trần cứng `min(tonKhoVatLy, conLaiCuaDon)`.
- Dùng cùng block chống kẹt kho.
- Dùng cùng cách tính tồn kho tự động từ tổng nhập/tổng xuất.
- Dùng cùng quy tắc làm tròn 4 chữ số cho số lượng.

Khác biệt duy nhất được phép:

- Map cột dữ liệu.
- Tên sheet.
- Quy tắc nghiệp vụ mở rộng độc lập (ví dụ: xuất theo cặp, ưu tiên vị trí...).

---

## 8) Pseudocode chuẩn (để tái sử dụng)

```ts
function round4(n) {
  return Number(Number(n || 0).toFixed(4));
}

function tinhSoLuongSoan(record, config) {
  const tongXuat = sum(record.dsXuatDaGhi);
  const tonKho = Math.max(0, round4(record.tongNhap - tongXuat));
  const chtThucTe = Math.max(0, round4(record.soLuongDonHang - record.tongNhap));
  let chtBang = Number(record.chuaHoanThanhBang);
  if (!isFinite(chtBang) || chtBang < 0) chtBang = chtThucTe;
  const chuaHT = Math.max(chtBang, chtThucTe);
  const nguong = getNguongTheoHopDong(record.hopDong, config);

  let conLaiDon = round4(record.soLuongDonHang - tongXuat);
  if (conLaiDon <= 0 && !record.laSKUBoSung) return skip("Đơn đã xuất đủ");
  if (conLaiDon < 0) conLaiDon = 0;

  let deXuat = Math.min(tonKho, conLaiDon);

  if (config.locTonKho && tonKho < nguong && chuaHT > 0 && !record.laSKUBoSung) {
    return skip("Tồn dưới ngưỡng, chưa hoàn thành > 0");
  }

  if (chuaHT > 0 && chuaHT < nguong) {
    const khaoSat = round4(tonKho + chuaHT - nguong);
    deXuat = Math.min(deXuat, khaoSat);
    if (deXuat < nguong && !record.laSKUBoSung) {
      return skip("Không đạt mức tối thiểu sau chống kẹt kho");
    }
  }

  deXuat = round4(deXuat);
  if (deXuat <= 0) return skip("SL gợi ý <= 0");
  return ok(deXuat);
}
```

---

## 9) Checklist triển khai khi map cột khác nhau

Khi áp dụng vào script mới:

1. Map cột vật lý -> schema logic chuẩn tại Mục 2.
2. Viết helper `readRecordFromRow(...)` trả về đúng schema.
3. Gọi chung hàm tính ở Mục 8.
4. Trả message lỗi theo ngữ cảnh script.
5. Bảo đảm test case tối thiểu:
   - Tồn kho tự tính lệch với cột tồn kho hiển thị.
   - CHT cột bảng âm/sai -> fallback sang CHT tự tính.
   - Kẹt kho (CHT > 0 và < ngưỡng) khiến SL gợi ý bị hạ.
   - SKU bổ sung có thể đi qua một số chặn ngưỡng.
   - Nhóm xuất theo cặp chưa đủ ngưỡng phải bị loại ở hậu xử lý.
