# Quy tắc ghép món giả lập hóa đơn

## Bối cảnh
Các đơn GrabMart tháng 1/2026 (file `BAO CAO T1 - XULY.xlsx`) không có PDF chi tiết món.
Tool dựng lại danh sách món cho từng đơn để xuất hóa đơn, dựa trên giá trong
`HÀNG BÁN CHẠY.xlsx` và `HÀNG BÁN KHÔNG CHẠY.xlsx` (cột **Đơn giá**).

## Dữ liệu đơn
- Ghép các đơn có trạng thái "Đã thanh toán" hoặc "Chưa thanh toán" (sau khi loại ký tự rác `_x0008_`, không phân biệt hoa thường). Trạng thái khác (Đã hủy, Thanh toán thất bại...) bị bỏ qua.
- Khóa đơn là **Mã đơn hàng dài** (mã ngắn GM-xxx bị trùng giữa các đơn).
- Bỏ qua (kèm lý do): thiếu mã đơn dài, số tiền trống/≤ 0, số tiền có phần lẻ, trùng mã đơn dài.
- Cảnh báo (vẫn ghép): Doanh thu ròng ≠ Số tiền + Giảm giá, thiếu doanh thu ròng, giảm giá dương, giảm giá lớn hơn số tiền.
- Tổng các dòng món phải **bằng đúng cột "Số tiền" (AD)**. Giảm giá (AJ) là dòng riêng, nên tổng sau giảm = Doanh thu ròng (AO).

## Danh mục món
- Tên món, tên cột và từ khóa loại trừ được làm sạch giống nhau: bỏ dấu `"` ở đầu, ký tự điều khiển, xuống dòng/tab → dấu cách, gộp khoảng trắng thừa/`\xa0`, chuẩn hóa Unicode NFC. So khớp tên không phân biệt hoa thường.
- Mỗi món lấy **một đơn giá**: giá nguyên dương có tổng số lượng bán lớn nhất. Giá có phần lẻ (VD 32.537,5) bị bỏ.
  Nếu ô Đơn giá trống thì dùng doanh thu ÷ số lượng.
- Tier 1 = món trong file bán chạy. Tier 2 = món trong file không chạy **mà chưa có** trong file bán chạy.
  Món có ở cả hai file dùng giá của file bán chạy.
- Loại món có tên chứa từ khóa loại trừ (mặc định `Hàng tặng`, tức `[Hàng tặng không bán] ...`).
- Trọng số chọn món = tổng số lượng bán (tier 1) hoặc số lượng được đặt (tier 2), tối thiểu 1.

## Quy tắc cứng (không bao giờ vi phạm)
1. Tổng thành tiền các dòng = Số tiền của đơn, chính xác đến từng đồng.
2. Đơn giá mỗi dòng = đơn giá của món trong danh mục; số lượng là số nguyên ≥ 1.
3. Một món không lặp lại 2 dòng trong cùng một đơn.
4. Số lượng mỗi dòng ≤ SL tối đa (mặc định 3; món rẻ dưới 20.000đ: 5).
5. **Trần tần suất**: một món có mặt trong tối đa `floor(cap% × số đơn đủ điều kiện)` đơn (mặc định 10%, tối thiểu 1 đơn).
6. Không dùng món bị loại trừ.

## Quy tắc mềm (hợp lý)
- Số dòng món theo số tiền: < 50k: 1–2 · 50k–<150k: 2–4 · 150k–<300k: 3–6 · ≥ 300k: 5–10.
  Khi "nới số món": từ `max(1, min−1)` đến `max + max(2, max//2)`.
- Món bán chạy được chọn nhiều hơn; món đã dùng nhiều bị giảm trọng số dần (xoay vòng).
- Hai đơn không có cùng tổ hợp (món, số lượng), trừ khi đơn được ghép ở bước "cho phép trùng tổ hợp".

## Thứ tự thử cho mỗi đơn (dừng ở bước đầu tiên thành công)
1. Bước 1 — chỉ hàng bán chạy, đúng khoảng số món.
2. Bước 1 (nới số món).
3. Bước 2 — thêm hàng không chạy, đúng khoảng số món.
4. Bước 2 (nới số món).
5. Bước 2 (cho phép trùng tổ hợp).
6. Không được → đánh dấu "Cần xử lý" kèm lý do.

Đơn được xử lý theo số tiền giảm dần (đơn lớn khó ghép hơn được ưu tiên trước khi các món chạm trần).
"% số đơn" của một món tính trên tổng số đơn đủ điều kiện (cùng mẫu số với trần).
Cùng seed + cùng dữ liệu + cùng tham số → cùng kết quả.

## Xuất file hóa đơn (mẫu nhập phần mềm V7)
Theo mẫu `Copy of MERCHANT_PORTAL_TONG_HOP_DA_XU_LY_02.xlsx`.

**Sheet "Hóa đơn bán hàng"**: dòng 1–6 tiêu đề + hướng dẫn, dòng 8 tên cột (B..S, cột A không có tên), dữ liệu từ dòng 9.
Mỗi hóa đơn là một khối, các khối cách nhau 1 dòng trống, xếp theo ngày tạo đơn tăng dần:
- Dòng món: A = mã GM (mã ngắn), B = số thứ tự hóa đơn (1, 2, 3...), C = ngày hóa đơn `dd/mm/yyyy`,
  G = "Bán cho người tiêu dùng", K = "TM/CK", L = tên món, N = số lượng, O = đơn giá, P = thành tiền,
  Q = "GRAB - <mã GM>".
- **1 dòng giảm giá** (in đậm): L = tên dòng giảm giá (mặc định "Giảm giá"), P = số tiền giảm
  (số dương, = −cột "Giảm giá (Quán tài trợ)"), R = 1. Đơn có giảm giá = 0 thì không có dòng này.
- Dòng tổng (in đậm, nền xanh #D9EAF7): L = "THÀNH TIỀN <mã GM>", P = tổng món − giảm giá (= Doanh thu ròng).
- Ngày hóa đơn: theo ngày tạo đơn (mặc định) hoặc một ngày cố định cho mọi hóa đơn.
- Mã GM bị trùng giữa các đơn, nên số thứ tự hóa đơn (cột B) mới là khóa phân biệt hóa đơn.
- Chỉ đơn ghép thành công mới có trong sheet này.

**Sheet "Kiểm tra"**: mỗi đơn một dòng — STT (= số thứ tự hóa đơn), Mã GM, Mã đơn hàng dài, Ngày tạo, Số dòng món,
Số dòng chiết khấu, Tổng SL, Tổng tạm tính (= Số tiền), Giảm giá thực tế, Tổng cộng (= Doanh thu ròng trên báo cáo),
Tổng tính lại (= dòng THÀNH TIỀN), Trạng thái (`OK` / `LỆCH ...` / `CHƯA GHÉP: <lý do>`). Đơn chưa ghép được nằm cuối, STT trống.
