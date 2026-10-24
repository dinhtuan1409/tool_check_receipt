"""Web giả lập hóa đơn cho các đơn GrabMart không có chi tiết món.

Chạy: streamlit run app.py
"""
import hashlib
from datetime import datetime
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

import invoice_engine as eng

DATA_DIR = Path(__file__).parent / "data"
DEFAULT_FILES = {
    "orders": "BAO CAO T1 - XULY.xlsx",
    "hot": "HÀNG BÁN CHẠY.xlsx",
    "cold": "HÀNG BÁN KHÔNG CHẠY.xlsx",
}
FILE_LABELS = {"orders": "Báo cáo đơn hàng", "hot": "Hàng bán chạy", "cold": "Hàng bán không chạy"}

st.set_page_config(page_title="Giả lập hóa đơn GrabMart", layout="wide")


def vnd(x) -> str:
    return f"{int(round(x)):,}".replace(",", ".")


# kết quả được cache theo nội dung file + tham số; spinner do nơi gọi hiển thị (kèm tên file)
@st.cache_data(show_spinner=False)
def load_data(orders_bytes, hot_bytes, cold_bytes, exclude):
    orders = eng.load_orders(orders_bytes)
    catalog, excluded = eng.build_catalog(eng.load_items(hot_bytes), eng.load_items(cold_bytes), list(exclude))
    return orders, catalog, excluded


@st.cache_data(show_spinner=False)
def run(orders, catalog, params: eng.Params):
    return (*eng.generate_invoices(orders, catalog, params), datetime.now())


def pick_file(key):
    up = st.sidebar.file_uploader(FILE_LABELS[key], type=["xlsx"], key=key)
    if up is not None:
        data, name, source = up.getvalue(), up.name, "Tải lên"
    elif (DATA_DIR / DEFAULT_FILES[key]).exists():
        data, name, source = (DATA_DIR / DEFAULT_FILES[key]).read_bytes(), DEFAULT_FILES[key], "Mặc định (thư mục data)"
    else:
        return None
    return {"bytes": data, "name": name, "source": source, "size": len(data),
            "digest": hashlib.md5(data).hexdigest()[:8]}


# ---------------------------------------------------------------- sidebar
st.sidebar.header("Dữ liệu")
files = {k: pick_file(k) for k in FILE_LABELS}

st.sidebar.header("Tham số")
seed = st.sidebar.number_input("Seed (cùng seed → cùng kết quả)", value=2026, step=1)
cap_pct = st.sidebar.slider("Trần: một món xuất hiện tối đa (% số đơn)", 1.0, 50.0, 10.0, 0.5)
max_qty = st.sidebar.number_input("SL tối đa mỗi dòng", 1, 20, 3)
cheap_price = st.sidebar.number_input("Món rẻ: đơn giá dưới (đ)", 0, 1_000_000, 20_000, step=1000)
max_qty_cheap = st.sidebar.number_input("SL tối đa mỗi dòng – món rẻ", 1, 20, 5)
exclude_raw = st.sidebar.text_input("Loại món có chứa (phân cách bằng dấu phẩy)", ", ".join(eng.DEFAULT_EXCLUDE))
exclude = tuple(k.strip() for k in exclude_raw.split(",") if k.strip())

st.title("Giả lập hóa đơn GrabMart")

missing = [FILE_LABELS[k] for k, f in files.items() if f is None]
if missing:
    st.info(f"Tải lên đủ 3 file Excel ở thanh bên trái để bắt đầu. Còn thiếu: {', '.join(missing)}.")
    st.stop()

params = eng.Params(seed=int(seed), cap_pct=float(cap_pct), max_qty=int(max_qty),
                    max_qty_cheap=int(max_qty_cheap), cheap_price=int(cheap_price))
names = ", ".join(f["name"] for f in files.values())
try:
    with st.spinner(f"Đang đọc file: {names} ..."):
        orders, catalog, excluded = load_data(files["orders"]["bytes"], files["hot"]["bytes"],
                                              files["cold"]["bytes"], exclude)
    with st.spinner(f"Đang ghép món cho {int(orders.eligible.sum()):,} đơn ..."):
        results, lines, cap, computed_at = run(orders, catalog, params)
except ValueError as err:
    st.error(str(err))
    st.stop()

# báo khi bộ file (hoặc tham số) khác lần trước -> người dùng biết kết quả đã được tính lại
digests = tuple(f["digest"] for f in files.values())
prev = st.session_state.get("signature")
if prev is not None and prev[0] != digests:
    st.toast(f"Đã tính xong cho file mới: {names}", icon="✅")
elif prev is not None and prev[1:] != (params, exclude):
    st.toast("Đã tính lại theo tham số mới", icon="✅")
st.session_state["signature"] = (digests, params, exclude)

with st.container(border=True):
    st.markdown("**Dữ liệu đang dùng**")
    for col, (k, f) in zip(st.columns(3), files.items()):
        col.markdown(f"{FILE_LABELS[k]}: `{f['name']}`")
        col.caption(f"{f['source']} · {vnd(f['size'] / 1024)} KB · mã nội dung {f['digest']}")
    st.caption(f"Kết quả tính lúc {computed_at:%H:%M:%S %d/%m/%Y} · {vnd(len(orders))} dòng trong báo cáo · "
               f"{vnd(len(results))} đơn cần ghép")

st.sidebar.header("Xuất hóa đơn")
discount_name = st.sidebar.text_input(
    "Tên dòng giảm giá", "Giảm giá",
    help="Mỗi hóa đơn có 1 dòng này, Thành tiền = số tiền giảm, cột Chiết khấu thương mại = 1. "
         "Đơn không có giảm giá thì không có dòng này.")
date_mode = st.sidebar.radio("Ngày hóa đơn", ["Theo ngày tạo đơn", "Một ngày cố định"])
invoice_date = None
if date_mode == "Một ngày cố định":
    invoice_date = pd.Timestamp(st.sidebar.date_input("Chọn ngày", format="DD/MM/YYYY"))

ok = results[results.ok]
failed = results[~results.ok]
invoice_no = eng.invoice_order(results).set_index("order_id").invoice_no
skipped = orders[~orders.eligible]
usage = eng.item_usage(lines, len(results))
problems = eng.validate(results, lines, catalog, params, cap, exclude)

tab_inv, tab_sum, tab_fail, tab_data = st.tabs(["Hóa đơn", "Tổng quan", "Cần xử lý", "Dữ liệu & tải về"])

# ---------------------------------------------------------------- hóa đơn
INVOICE_CSS = """
<style>
.inv{background:#fff;color:#1b1b1b;max-width:760px;margin:8px auto;padding:28px 32px;border:1px solid #d9d9d9;
     border-radius:6px;font-family:Arial,Helvetica,sans-serif;font-size:14px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.inv h2{margin:0;font-size:20px;letter-spacing:.5px;text-align:center}
.inv .sub{text-align:center;color:#c0392b;font-size:12px;margin:2px 0 16px}
.inv .shop{display:flex;justify-content:space-between;border-bottom:2px solid #1b1b1b;padding-bottom:10px;margin-bottom:12px}
.inv .shop b{font-size:16px}
.inv .meta{display:grid;grid-template-columns:150px 1fr;row-gap:3px;margin-bottom:14px}
.inv .meta span:nth-child(odd){color:#666}
.inv table{width:100%;border-collapse:collapse}
.inv th{background:#f3f3f3;text-align:left;padding:7px 6px;border-bottom:1px solid #ccc;font-size:13px}
.inv td{padding:6px;border-bottom:1px solid #eee;vertical-align:top}
.inv .r{text-align:right;white-space:nowrap}
.inv tr.disc td{font-weight:bold;color:#a0522d}
.inv .tag{font-size:11px;color:#8a6d00;background:#fff4cc;border-radius:3px;padding:0 4px;margin-left:4px}
.inv .tot{margin-top:12px;margin-left:auto;width:340px}
.inv .tot div{display:flex;justify-content:space-between;padding:3px 0}
.inv .tot .grand{border-top:2px solid #1b1b1b;margin-top:4px;padding-top:6px;font-weight:bold;font-size:16px}
.inv .foot{margin-top:18px;color:#888;font-size:12px;text-align:center}
</style>
"""


def invoice_html(o, items: pd.DataFrame) -> str:
    net = items.total.sum() + o.discount
    disc_row = ("" if o.discount == 0 else
                f"<tr class=disc><td></td><td>{escape(discount_name)}</td><td></td><td></td>"
                f"<td class=r>-{vnd(-o.discount)}</td></tr>")
    net_note = ("" if pd.isna(o.net) or abs(o.net - net) <= 0.5 else
                f"<div style='color:#c0392b'><span>Doanh thu ròng trên báo cáo</span><span>{vnd(o.net)} đ</span></div>")
    rows = "".join(
        f"<tr><td>{i}</td><td>{escape(r.name)}"
        f"{'<span class=tag>không chạy</span>' if r.tier == 2 else ''}</td>"
        f"<td class=r>{r.qty}</td><td class=r>{vnd(r.price)}</td><td class=r>{vnd(r.total)}</td></tr>"
        for i, r in enumerate(items.itertuples(), 1))
    return f"""{INVOICE_CSS}
<div class="inv">
  <div class="shop"><div><b>{escape(o.store or 'Cửa hàng')}</b><br>{escape(o.service or 'GrabMart')}</div>
    <div class="r">Đơn: <b>{escape(o.short_id)}</b><br>{escape(o.created)}</div></div>
  <h2>HÓA ĐƠN BÁN HÀNG</h2>
  <div class="sub">Giả lập — chi tiết món được ghép tự động</div>
  <div class="meta">
    <span>Số thứ tự hóa đơn</span><span>{invoice_no.get(o.order_id, '')}</span>
    <span>Mã đơn hàng dài</span><span>{escape(o.order_id)}</span>
    <span>Mã đặt hàng</span><span>{escape(o.booking_id)}</span>
    <span>Thanh toán</span><span>{escape(o.payment)}</span>
  </div>
  <table><thead><tr><th>#</th><th>Tên hàng</th><th class=r>SL</th><th class=r>Đơn giá</th><th class=r>Thành tiền</th></tr></thead>
  <tbody>{rows}{disc_row}</tbody></table>
  <div class="tot">
    <div><span>Tổng tiền hàng</span><span>{vnd(items.total.sum())}</span></div>
    <div><span>Giảm giá</span><span>-{vnd(-o.discount)}</span></div>
    <div class="grand"><span>Thành tiền</span><span>{vnd(net)} đ</span></div>
    {net_note}
  </div>
  <div class="foot">{len(items)} dòng món · {int(items.qty.sum())} sản phẩm · {escape(o.step)}</div>
</div>"""


with tab_inv:
    c1, c2 = st.columns([2, 1])
    query = c1.text_input("Tìm theo mã đơn dài / mã đơn ngắn / mã đặt hàng", placeholder="VD: GM-123")
    show = c2.radio("Hiển thị", ["Tất cả", "Đã ghép được", "Cần xử lý"], horizontal=True)
    view = results
    if show == "Đã ghép được":
        view = ok
    elif show == "Cần xử lý":
        view = failed
    if query.strip():
        q = query.strip().casefold()
        view = view[view.order_id.str.casefold().str.contains(q, regex=False)
                    | view.short_id.str.casefold().str.contains(q, regex=False)
                    | view.booking_id.str.casefold().str.contains(q, regex=False)]
    if view.empty:
        st.warning("Không có đơn nào phù hợp.")
    else:
        labels = {r.order_id: f"{r.short_id} · {r.order_id} · {vnd(r.amount)}đ{'' if r.ok else ' · LỖI'}"
                  for r in view.itertuples()}
        oid = st.selectbox(f"Chọn đơn ({len(view):,} đơn)", list(labels), format_func=labels.get)
        o = view[view.order_id == oid].iloc[0]
        if o.warning:
            st.warning(o.warning)
        if not o.ok:
            st.error(f"Đơn {o.short_id} ({vnd(o.amount)}đ) chưa ghép được: {o.fail_reason}")
        else:
            st.html(invoice_html(o, lines[lines.order_id == oid]))

# ---------------------------------------------------------------- tổng quan
with tab_sum:
    m = st.columns(5)
    m[0].metric("Đơn cần ghép (đã + chưa thanh toán)", vnd(len(results)))
    m[1].metric("Ghép được", vnd(len(ok)))
    m[2].metric("Cần xử lý", vnd(len(failed)))
    m[3].metric("Bỏ qua (hủy, lỗi dữ liệu)", vnd(len(skipped)))
    top_pct = usage["% số đơn"].max() if not usage.empty else 0
    m[4].metric("Món phổ biến nhất", f"{top_pct:.1f}% số đơn", help=f"Trần: {cap:,} đơn ({cap_pct:g}%)")

    if problems:
        st.error(f"Kiểm tra tự động: {len(problems)} vi phạm")
        st.write(problems[:50])
    else:
        st.success("Kiểm tra tự động: mọi hóa đơn khớp đúng số tiền, đúng đơn giá, "
                   "không món nào vượt trần, không trùng tổ hợp ngoài các đơn được phép.")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Đơn theo bước ghép")
        st.dataframe(ok.step.value_counts().rename_axis("Bước").reset_index(name="Số đơn"), hide_index=True)
        st.subheader("Số dòng món mỗi đơn")
        st.bar_chart(ok.n_lines.value_counts().sort_index().rename("Số đơn"))
    with c2:
        st.subheader("Danh mục món")
        st.write(f"{(catalog.tier == 1).sum():,} món bán chạy · {(catalog.tier == 2).sum():,} món chỉ có ở file "
                 f"không chạy · {len(excluded):,} món bị loại")
        if not excluded.empty:
            st.dataframe(excluded, hide_index=True)

    st.subheader("Tần suất xuất hiện của từng món")
    st.dataframe(usage, hide_index=True, height=420, column_config={
        "% số đơn": st.column_config.ProgressColumn("% số đơn", format="%.2f%%", min_value=0,
                                                   max_value=max(float(cap_pct), top_pct)),
        "Đơn giá": st.column_config.NumberColumn(format="localized"),
        "Thành tiền": st.column_config.NumberColumn(format="localized"),
    })

# ---------------------------------------------------------------- cần xử lý
with tab_fail:
    st.subheader(f"Đơn chưa ghép được ({len(failed):,})")
    if failed.empty:
        st.success("Không có đơn nào.")
    else:
        st.dataframe(failed[["short_id", "order_id", "created", "amount", "fail_reason"]].rename(columns={
            "short_id": "Mã ngắn", "order_id": "Mã đơn dài", "created": "Ngày tạo",
            "amount": "Số tiền", "fail_reason": "Lý do"}), hide_index=True)
    warned = results[results.warning != ""]
    st.subheader(f"Đơn có cảnh báo dữ liệu ({len(warned):,})")
    if not warned.empty:
        st.dataframe(warned[["short_id", "order_id", "amount", "discount", "net", "warning"]], hide_index=True)
    st.subheader(f"Dòng bị bỏ qua trong file báo cáo ({len(skipped):,})")
    if not skipped.empty:
        st.dataframe(skipped[["row", "short_id", "order_id", "status", "amount", "skip_reason"]].rename(columns={
            "row": "Dòng Excel", "short_id": "Mã ngắn", "order_id": "Mã đơn dài", "status": "Trạng thái",
            "amount": "Số tiền", "skip_reason": "Lý do"}), hide_index=True)

# ---------------------------------------------------------------- dữ liệu
with tab_data:
    xlsx_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.download_button(f"Tải file hóa đơn – mẫu nhập phần mềm V7 ({len(ok):,} hóa đơn)",
                       data=lambda: eng.to_template_excel(results, lines, discount_name, invoice_date),
                       file_name="hoa_don_nhap_phan_mem.xlsx", type="primary", mime=xlsx_mime)
    st.caption("Sheet 'Hóa đơn bán hàng': dòng món → 1 dòng giảm giá → dòng THÀNH TIỀN, theo thứ tự ngày tạo đơn. "
               "Sheet 'Kiểm tra': đối chiếu từng hóa đơn với doanh thu ròng trên báo cáo.")
    st.subheader(f"Chi tiết dòng món ({len(lines):,} dòng)")
    st.dataframe(lines.merge(results[["order_id", "short_id"]], on="order_id"), hide_index=True, height=480)
