"""Ghép món giả lập cho các đơn GrabMart không có chi tiết, để dựng hóa đơn.

Quy tắc đầy đủ: docs/quy-tac-ghep-mon.md
"""
from __future__ import annotations

import io
import math
import re
import unicodedata
from dataclasses import dataclass

import numpy as np
import pandas as pd

# (giới hạn trên của số tiền, số dòng món tối thiểu, tối đa) — áp dụng khi số tiền < giới hạn
BRACKETS = [(50_000, 1, 2), (150_000, 2, 4), (300_000, 3, 6), (math.inf, 5, 10)]
# xác suất chọn số lượng 1, 2, 3, 4, 5 cho một dòng món; từ 6 trở lên dùng QTY_PROB_TAIL
QTY_PROBS = [0.65, 0.25, 0.07, 0.02, 0.01]
QTY_PROB_TAIL = 0.005
# độ ưu tiên món có giá gần "giá mục tiêu" của dòng (số tiền còn lại / số dòng còn lại)
CLOSENESS_ALPHA = 1.0
# trọng số tối thiểu của món đã dùng gần chạm trần (để xoay vòng món)
MIN_USAGE_FACTOR = 0.05
# số tổ hợp tối đa gom lại khi tìm món lấp tiền
MAX_FILL_OPTIONS = 60
DEFAULT_EXCLUDE = ["Hàng tặng"]
# trạng thái đơn được ghép món (so sánh sau clean_text + casefold)
INCLUDED_STATUSES = {"đã thanh toán", "chưa thanh toán"}

# (nhãn, các tier được dùng, nới số món, cho phép trùng tổ hợp)
STEPS = [
    ("Bước 1", frozenset({1}), False, False),
    ("Bước 1 (nới số món)", frozenset({1}), True, False),
    ("Bước 2", frozenset({1, 2}), False, False),
    ("Bước 2 (nới số món)", frozenset({1, 2}), True, False),
    ("Bước 2 (cho phép trùng tổ hợp)", frozenset({1, 2}), True, True),
]


@dataclass(frozen=True)
class Params:
    seed: int = 2026
    cap_pct: float = 10.0
    max_qty: int = 3
    max_qty_cheap: int = 5
    cheap_price: int = 20_000
    attempts: int = 80


# ---------------------------------------------------------------- đọc dữ liệu

def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    s = str(value).replace("_x0008_", "").replace("\xa0", " ")
    s = re.sub(r"[\t\n\r\v\f]", " ", s)  # xuống dòng (Alt+Enter) / tab -> dấu cách, không dính chữ
    s = "".join(ch for ch in s if not unicodedata.category(ch).startswith("C"))
    s = unicodedata.normalize("NFC", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lstrip('"').strip()


def _as_io(src):
    return io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src


def _read_sheet(src) -> pd.DataFrame:
    raw = pd.read_excel(_as_io(src), sheet_name=0)
    raw.columns = [clean_text(c) for c in raw.columns]
    return raw


def _first_col(df: pd.DataFrame, *names: str) -> str | None:
    return next((n for n in names if n in df.columns), None)


def _is_int_price(p) -> bool:
    return pd.notna(p) and p >= 1 and abs(p - round(p)) < 1e-6


def normalize_keywords(keywords) -> list[str]:
    return [k for k in (clean_text(k).casefold() for k in keywords or []) if k]


def parse_created(values: pd.Series) -> pd.Series:
    """'31 Jan 2026 9:17 PM' -> datetime; định dạng khác thử đọc kiểu ngày/tháng/năm."""
    dt = pd.to_datetime(values, format="%d %b %Y %I:%M %p", errors="coerce")
    rest = dt.isna() & (values != "")
    if rest.any():
        dt[rest] = pd.to_datetime(values[rest], dayfirst=True, errors="coerce", format="mixed")
    return dt


def load_orders(src) -> pd.DataFrame:
    """Đọc file báo cáo đơn; đánh dấu đơn nào đủ điều kiện ghép món."""
    raw = _read_sheet(src)
    missing = [c for c in ("Mã đơn hàng dài", "Trạng thái", "Số tiền") if c not in raw.columns]
    if missing:
        raise ValueError(f"File báo cáo thiếu cột: {', '.join(missing)}")

    def col(name):
        return raw[name] if name in raw.columns else pd.Series([None] * len(raw), index=raw.index)

    df = pd.DataFrame({
        "row": raw.index + 2,
        "order_id": col("Mã đơn hàng dài").map(clean_text),
        "short_id": col("Mã đơn hàng ngắn").map(clean_text),
        "booking_id": col("Mã đặt hàng").map(clean_text),
        "store": col("Tên cửa hàng").map(clean_text),
        "service": col("Loại").map(clean_text),
        "created": col("Ngày tạo").map(clean_text),
        "status": col("Trạng thái").map(clean_text),
        "payment": col("Phương thức thanh toán").map(clean_text),
        "amount": pd.to_numeric(col("Số tiền"), errors="coerce"),
        "discount": pd.to_numeric(col("Giảm giá (Quán tài trợ)"), errors="coerce").fillna(0),
        "net": pd.to_numeric(col("Doanh thu ròng"), errors="coerce"),
    })

    eligible, skip, warn = [], [], []
    seen: dict[str, int] = {}
    for r in df.itertuples():
        reason = ""
        if r.status.casefold() not in INCLUDED_STATUSES:
            reason = f"Trạng thái: {r.status or '(trống)'}"
        elif not r.order_id:
            reason = "Thiếu mã đơn hàng dài"
        elif pd.isna(r.amount) or r.amount <= 0:
            reason = "Số tiền trống hoặc ≤ 0"
        elif not _is_int_price(r.amount):
            reason = "Số tiền có phần lẻ"
        elif r.order_id in seen:
            reason = f"Trùng mã đơn với dòng {seen[r.order_id]}"
        else:
            seen[r.order_id] = r.row

        notes = []
        if not reason:
            if pd.isna(r.net):
                notes.append("Thiếu doanh thu ròng")
            elif abs(r.net - (r.amount + r.discount)) > 0.5:
                notes.append("Doanh thu ròng ≠ Số tiền + Giảm giá")
            if r.discount > 0:
                notes.append("Giảm giá là số dương")
            if abs(r.discount) > r.amount:
                notes.append("Giảm giá lớn hơn số tiền")
        eligible.append(not reason)
        skip.append(reason)
        warn.append("; ".join(notes))

    df["eligible"] = eligible
    df["skip_reason"] = skip
    df["warning"] = warn
    df["created_dt"] = parse_created(df.created)
    df.loc[df.eligible, "amount"] = df.loc[df.eligible, "amount"].round()
    return df


def load_items(src) -> pd.DataFrame:
    """Đọc file hàng bán chạy hoặc hàng bán không chạy về dạng: name, price, weight."""
    raw = _read_sheet(src)
    name_c = _first_col(raw, "Item", "Mục", "Sản phẩm")
    price_c = _first_col(raw, "Đơn giá")
    qty_c = _first_col(raw, "Units Sold", "Số lượng sản phẩm đã bán", "Số lượng thiếu hụt")
    amt_c = _first_col(raw, "Item Gross Sales (₫)", "Doanh thu gộp của mặt hàng (₫)", "Doanh thu bị mất (₫)")
    weight_c = _first_col(raw, "Units Sold", "Số lượng sản phẩm đã bán", "Số lượng được đặt")
    if not name_c or not (price_c or (qty_c and amt_c)):
        raise ValueError("File mặt hàng cần có cột tên món và cột 'Đơn giá' (hoặc số lượng + doanh thu)")

    def num(c):
        return pd.to_numeric(raw[c], errors="coerce") if c else pd.Series(np.nan, index=raw.index)

    qty, amount = num(qty_c), num(amt_c)
    price = num(price_c)
    fallback = price.isna() & (qty > 0)
    price[fallback] = amount[fallback] / qty[fallback]
    df = pd.DataFrame({
        "name": raw[name_c].map(clean_text),
        "price": price,
        "weight": num(weight_c).fillna(0).clip(lower=0),
    })
    # dòng tiêu đề tiếng Việt lặp lại (Ngày, Mục, ...) có đơn giá không phải số -> bị loại ở đây
    return df[(df.name != "") & df.price.notna()].reset_index(drop=True)


def aggregate_items(items: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple[str, str]]]:
    """Gộp theo tên món; mỗi món lấy đơn giá nguyên phổ biến nhất (theo số lượng bán)."""
    rows, dropped = [], []
    items = items.assign(key=items.name.str.casefold())
    for key, g in items.groupby("key", sort=False):
        name = g.name.value_counts().index[0]
        ok = g[g.price.map(_is_int_price)]
        if ok.empty:
            dropped.append((name, "Không có đơn giá nguyên dương"))
            continue
        by_price = (ok.assign(p=ok.price.round().astype(int))
                    .groupby("p").agg(w=("weight", "sum"), n=("price", "size"))
                    .reset_index()
                    .sort_values(["w", "n", "p"], ascending=False))
        rows.append({
            "key": key, "name": name, "price": int(by_price.p.iat[0]),
            "weight": float(g.weight.sum()), "n_prices": len(by_price),
        })
    return pd.DataFrame(rows, columns=["key", "name", "price", "weight", "n_prices"]), dropped


def build_catalog(hot: pd.DataFrame, cold: pd.DataFrame, exclude_keywords=DEFAULT_EXCLUDE):
    """Tier 1 = hàng bán chạy; tier 2 = hàng bán không chạy chưa có trong tier 1."""
    hot_agg, hot_drop = aggregate_items(hot)
    cold_agg, cold_drop = aggregate_items(cold)
    hot_agg["tier"] = 1
    cold_agg = cold_agg[~cold_agg.key.isin(hot_agg.key)].assign(tier=2)
    catalog = pd.concat([hot_agg, cold_agg], ignore_index=True)

    kws = normalize_keywords(exclude_keywords)
    mask = catalog.key.map(lambda k: any(kw in k for kw in kws)).astype(bool)
    excluded = [(n, "Chứa từ khóa loại trừ") for n in catalog.loc[mask, "name"]]
    # món bị bỏ ở file này nhưng vẫn có giá hợp lệ ở file kia thì không tính là bị loại
    seen, kept = set(), set(catalog.key)
    for n, r in hot_drop + cold_drop:
        if n.casefold() not in kept and n.casefold() not in seen:
            seen.add(n.casefold())
            excluded.append((n, r))
    catalog = catalog[~mask].reset_index(drop=True)
    catalog["weight"] = catalog.weight.clip(lower=1)
    return catalog, pd.DataFrame(excluded, columns=["Món", "Lý do loại"])


# ---------------------------------------------------------------- ghép món

def line_bounds(amount: float) -> tuple[int, int]:
    for limit, lo, hi in BRACKETS:
        if amount < limit:
            return lo, hi
    raise AssertionError("unreachable")


def relaxed_bounds(lo: int, hi: int) -> tuple[int, int]:
    return max(1, lo - 1), hi + max(2, hi // 2)


def _closeness(values, target):
    return np.exp(-CLOSENESS_ALPHA * np.abs(np.log(np.asarray(values, dtype=float) / target)))


def _qty_prob(q: int) -> float:
    return QTY_PROBS[q - 1] if q <= len(QTY_PROBS) else QTY_PROB_TAIL


class InvoiceGenerator:
    def __init__(self, catalog: pd.DataFrame, params: Params, n_orders: int):
        if catalog.empty:
            raise ValueError("Không còn món nào để ghép")
        self.p = params
        self.names = catalog.name.tolist()
        self.prices = catalog.price.to_numpy(dtype=np.int64)
        self.weights = catalog.weight.to_numpy(dtype=float)
        self.tiers = catalog.tier.to_numpy()
        self.max_q = np.where(self.prices < params.cheap_price, params.max_qty_cheap, params.max_qty)
        # trần số đơn chứa một món: không vượt cap_pct% tổng số đơn (tối thiểu 1)
        self.cap = max(1, math.floor(params.cap_pct / 100 * n_orders + 1e-9))
        self.usage = np.zeros(len(catalog), dtype=np.int64)
        self.signatures: set[frozenset] = set()
        self.rng = np.random.default_rng(params.seed)
        self._value_maps: dict[frozenset, tuple[dict, list]] = {}

    # trọng số hiệu lực: bán chạy hơn -> hay được chọn hơn; dùng nhiều rồi -> giảm dần
    def _eff_weight(self, idx):
        return self.weights[idx] * np.maximum(MIN_USAGE_FACTOR, 1 - self.usage[idx] / self.cap)

    def _value_map(self, tiers):
        """giá trị dòng (đơn giá × SL) -> [(món, SL)], dùng để tìm món lấp tiền."""
        if tiers not in self._value_maps:
            vmap: dict[int, list] = {}
            for idx in np.flatnonzero(np.isin(self.tiers, list(tiers))):
                for q in range(1, int(self.max_q[idx]) + 1):
                    vmap.setdefault(int(self.prices[idx]) * q, []).append((int(idx), q))
            self._value_maps[tiers] = (vmap, sorted(vmap))
        return self._value_maps[tiers]

    def generate(self, amount: int):
        """Trả về ({idx: qty}, nhãn bước) hoặc (None, lý do)."""
        lo, hi = line_bounds(amount)
        for label, tiers, relax, allow_dup in STEPS:
            l_lo, l_hi = relaxed_bounds(lo, hi) if relax else (lo, hi)
            for _ in range(self.p.attempts):
                lines = self._attempt(amount, tiers, l_lo, l_hi, allow_dup)
                if lines:
                    self.usage[list(lines)] += 1
                    self.signatures.add(frozenset(lines.items()))
                    return lines, label
        avail = (self.usage < self.cap)
        if not avail.any():
            return None, "Mọi món đều đã chạm trần tần suất"
        if amount < self.prices[avail].min():
            return None, "Số tiền nhỏ hơn giá món rẻ nhất còn dùng được"
        return None, "Không tìm được tổ hợp món khớp đúng số tiền"

    def _attempt(self, amount, tiers, lo, hi, allow_dup):
        rng = self.rng
        pool = np.isin(self.tiers, list(tiers)) & (self.usage < self.cap)
        if not pool.any():
            return None
        min_price = int(self.prices[pool].min())
        n_lines = int(rng.integers(lo, hi + 1))
        n_fill = 1 if n_lines == 1 else int(rng.choice([1, 2, 2]))

        lines: dict[int, int] = {}
        remaining = amount
        for k in range(n_lines - n_fill):
            left = n_lines - k
            budget = remaining - min_price * (left - 1)
            target = remaining / left
            mask = pool & (self.prices <= budget)
            mask[list(lines)] = False
            cands = np.flatnonzero(mask)
            if cands.size == 0:
                return None
            w = self._eff_weight(cands) * _closeness(self.prices[cands], target)
            idx = int(rng.choice(cands, p=w / w.sum()))
            q = self._pick_qty(idx, budget, target)
            lines[idx] = q
            remaining -= int(self.prices[idx]) * q

        fill = self._exact_fill(remaining, n_fill, tiers, lines, lo, hi)
        if fill is None:
            return None
        lines.update(fill)
        if not allow_dup and frozenset(lines.items()) in self.signatures:
            return None
        return lines

    def _pick_qty(self, idx, budget, target):
        max_q = int(min(self.max_q[idx], budget // self.prices[idx]))
        qs = np.arange(1, max_q + 1)
        w = np.array([_qty_prob(q) for q in qs]) * _closeness(self.prices[idx] * qs, target)
        return int(self.rng.choice(qs, p=w / w.sum()))

    def _exact_fill(self, remaining, n_fill, tiers, used, lo, hi):
        """Tìm 1–2 dòng món (chưa có trong đơn, chưa chạm trần) cộng lại đúng `remaining`."""
        if remaining <= 0:
            return None
        vmap, values = self._value_map(tiers)

        def ok(u):
            return u[0] not in used and self.usage[u[0]] < self.cap

        counts = [n_fill] + [n for n in (1, 2) if n != n_fill and lo <= len(used) + n <= hi]
        for n in counts:
            options = []
            if n == 1:
                options = [(u,) for u in vmap.get(remaining, ()) if ok(u)]
            else:
                pairs = [v for v in values if v * 2 <= remaining and (remaining - v) in vmap]
                for i in self.rng.permutation(len(pairs)):
                    v = pairs[i]
                    for a in vmap[v]:
                        if not ok(a):
                            continue
                        for b in vmap[remaining - v]:
                            if b[0] != a[0] and ok(b) and (v * 2 < remaining or a < b):
                                options.append((a, b))
                    if len(options) >= MAX_FILL_OPTIONS:
                        break
            if options:
                w = np.array([np.prod([self._eff_weight(u[0]) * _qty_prob(u[1]) for u in opt]) for opt in options])
                opt = options[int(self.rng.choice(len(options), p=w / w.sum()))]
                return dict(opt)
        return None


def generate_invoices(orders: pd.DataFrame, catalog: pd.DataFrame, params: Params):
    """Ghép món cho mọi đơn đủ điều kiện. Trả về (kết quả từng đơn, chi tiết dòng món)."""
    todo = orders[orders.eligible].sort_values(["amount", "order_id"], ascending=[False, True])
    gen = InvoiceGenerator(catalog, params, len(todo))
    results, lines_out = [], []
    for o in todo.itertuples():
        lines, info = gen.generate(int(o.amount))
        ok = lines is not None
        results.append({
            "order_id": o.order_id, "ok": ok,
            "step": info if ok else "", "fail_reason": "" if ok else info,
            "n_lines": len(lines) if ok else 0,
            "n_units": sum(lines.values()) if ok else 0,
            "items_total": sum(int(gen.prices[i]) * q for i, q in lines.items()) if ok else 0,
        })
        if ok:
            for idx, q in sorted(lines.items(), key=lambda kv: gen.names[kv[0]]):
                price = int(gen.prices[idx])
                lines_out.append({
                    "order_id": o.order_id, "name": gen.names[idx], "tier": int(gen.tiers[idx]),
                    "price": price, "qty": q, "total": price * q,
                })
    res = pd.DataFrame(results, columns=["order_id", "ok", "step", "fail_reason", "n_lines", "n_units", "items_total"])
    res = orders[orders.eligible].merge(res, on="order_id", how="left")
    res["ok"] = res["ok"].fillna(False).astype(bool)
    lines_df = pd.DataFrame(lines_out, columns=["order_id", "name", "tier", "price", "qty", "total"])
    return res, lines_df, gen.cap


# ---------------------------------------------------------------- thống kê & kiểm tra

def item_usage(lines: pd.DataFrame, n_orders: int) -> pd.DataFrame:
    if lines.empty:
        return pd.DataFrame(columns=["Món", "Nhóm", "Đơn giá", "Số đơn", "% số đơn", "Tổng SL", "Thành tiền"])
    u = (lines.groupby(["name", "tier", "price"])
         .agg(n_orders=("order_id", "nunique"), qty=("qty", "sum"), total=("total", "sum"))
         .reset_index().sort_values("n_orders", ascending=False))
    u["pct"] = u.n_orders / max(n_orders, 1) * 100
    u["tier"] = u.tier.map({1: "Bán chạy", 2: "Không chạy"})
    return u.rename(columns={"name": "Món", "tier": "Nhóm", "price": "Đơn giá", "n_orders": "Số đơn",
                             "pct": "% số đơn", "qty": "Tổng SL", "total": "Thành tiền"})[
        ["Món", "Nhóm", "Đơn giá", "Số đơn", "% số đơn", "Tổng SL", "Thành tiền"]]


def validate(results: pd.DataFrame, lines: pd.DataFrame, catalog: pd.DataFrame,
             params: Params, cap: int, exclude_keywords=DEFAULT_EXCLUDE) -> list[str]:
    """Kiểm tra lại toàn bộ kết quả theo các quy tắc cứng. Danh sách rỗng = đạt."""
    problems = []
    ok = results[results.ok == True]  # noqa: E712
    bad_total = lines[lines.total != lines.price * lines.qty]
    for r in bad_total.itertuples():
        problems.append(f"Đơn {r.order_id}: '{r.name}' thành tiền {r.total:,} ≠ {r.price:,} × {r.qty}")
    orphan = sorted(set(lines.order_id) - set(ok.order_id))
    for oid in orphan:
        problems.append(f"Có dòng món gắn với đơn {oid} không phải đơn ghép thành công")
    sums = lines.groupby("order_id").total.sum()
    for o in ok.itertuples():
        if sums.get(o.order_id, 0) != int(o.amount):
            problems.append(f"Đơn {o.order_id}: tổng món {sums.get(o.order_id, 0):,} ≠ số tiền {int(o.amount):,}")
    price_of = dict(zip(catalog.name, catalog.price))
    bad_price = lines[lines.apply(lambda r: price_of.get(r["name"]) != r["price"], axis=1)]
    for r in bad_price.itertuples():
        problems.append(f"Đơn {r.order_id}: '{r.name}' giá {r.price:,} không khớp danh mục")
    if (lines.qty < 1).any():
        problems.append("Có dòng số lượng < 1")
    if (lines.qty != lines.qty.round()).any():
        problems.append("Có dòng số lượng không phải số nguyên")
    cheap = lines.price < params.cheap_price
    over = lines[(cheap & (lines.qty > params.max_qty_cheap)) | (~cheap & (lines.qty > params.max_qty))]
    for r in over.itertuples():
        problems.append(f"Đơn {r.order_id}: '{r.name}' số lượng {r.qty} vượt mức tối đa")
    dup_lines = lines.duplicated(["order_id", "name"])
    if dup_lines.any():
        problems.append(f"{int(dup_lines.sum())} dòng món bị lặp trong cùng một đơn")
    n_orders = lines.groupby("name").order_id.nunique()
    for name, n in n_orders[n_orders > cap].items():
        problems.append(f"Món '{name}' có trong {n} đơn, vượt trần {cap}")
    kws = normalize_keywords(exclude_keywords)
    for name in lines.name.unique():
        if any(kw in name.casefold() for kw in kws):
            problems.append(f"Món bị loại trừ vẫn xuất hiện: '{name}'")
    n_lines = lines.groupby("order_id").size()
    for o in ok.itertuples():
        lo, hi = line_bounds(o.amount)
        if "nới" in o.step or "trùng" in o.step:
            lo, hi = relaxed_bounds(lo, hi)
        n = int(n_lines.get(o.order_id, 0))
        if o.n_lines != n:
            problems.append(f"Đơn {o.order_id}: ghi nhận {o.n_lines} dòng món nhưng thực tế có {n}")
        if not (lo <= n <= hi):
            problems.append(f"Đơn {o.order_id}: {n} dòng món ngoài khoảng {lo}–{hi}")
    sigs = lines.groupby("order_id").apply(
        lambda g: tuple(sorted(zip(g["name"], g["qty"]))), include_groups=False)
    allowed_dup = set(ok.loc[ok.step.str.contains("trùng"), "order_id"])
    dup_sig = sigs[sigs.duplicated(keep="first") & ~sigs.index.isin(allowed_dup)]
    for oid in dup_sig.index:
        problems.append(f"Đơn {oid}: trùng tổ hợp món với đơn khác")
    return problems


def to_excel_bytes(results: pd.DataFrame, lines: pd.DataFrame, usage: pd.DataFrame) -> bytes:
    head = results.rename(columns={
        "order_id": "Mã đơn hàng dài", "short_id": "Mã đơn hàng ngắn", "booking_id": "Mã đặt hàng",
        "created": "Ngày tạo", "amount": "Số tiền", "discount": "Giảm giá (Quán tài trợ)",
        "net": "Doanh thu ròng", "step": "Bước ghép", "fail_reason": "Lý do lỗi", "warning": "Cảnh báo",
        "n_lines": "Số dòng món", "n_units": "Tổng SL",
    })[["Mã đơn hàng dài", "Mã đơn hàng ngắn", "Mã đặt hàng", "Ngày tạo", "Số tiền",
        "Giảm giá (Quán tài trợ)", "Doanh thu ròng", "Số dòng món", "Tổng SL", "Bước ghép", "Lý do lỗi", "Cảnh báo"]]
    detail = lines.merge(results[["order_id", "short_id", "created"]], on="order_id", how="left")
    detail["tier"] = detail.tier.map({1: "Bán chạy", 2: "Không chạy"})
    detail = detail.rename(columns={
        "order_id": "Mã đơn hàng dài", "short_id": "Mã đơn hàng ngắn", "created": "Ngày tạo",
        "name": "Tên hàng", "tier": "Nhóm", "price": "Đơn giá", "qty": "Số lượng", "total": "Thành tiền",
    })[["Mã đơn hàng dài", "Mã đơn hàng ngắn", "Ngày tạo", "Tên hàng", "Nhóm", "Đơn giá", "Số lượng", "Thành tiền"]]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        head.to_excel(xw, sheet_name="Don hang", index=False)
        detail.to_excel(xw, sheet_name="Chi tiet mon", index=False)
        usage.to_excel(xw, sheet_name="Tan suat mon", index=False)
    return buf.getvalue()


# ---------------------------------------------------------------- xuất theo mẫu phần mềm hóa đơn (V7)

TEMPLATE_TITLE = "File mẫu danh sách hóa đơn để nhập vào phần mềm - V7"
TEMPLATE_GUIDE = [
    "Hướng dẫn:",
    "- Điền dữ liệu hóa đơn cần lập trên phần mềm vào các cột tương ứng trên file này",
    "- Các cột có dấu (*) là những cột bắt buộc",
    "- Nếu muốn nhập thêm thông tin khác, người dùng có thể tự thêm cột trên file này "
    "(VD: Mã khách hàng, Mã hàng, Tỷ lệ chiết khấu)",
    "- Các dòng dữ liệu phía dưới chỉ là ví dụ minh họa",
]
# cột B..S của dòng tiêu đề (dòng 8); cột A không có tiêu đề, chứa mã GM
TEMPLATE_HEADERS = [
    "Số thứ tự hóa đơn (*)", "Ngày hóa đơn", "Tên đơn vị mua hàng", "Mã số thuế", "Địa chỉ",
    "Người mua hàng", "Email", "Số điện thoại", "Căn cước công dân", "Hình thức thanh toán (*)",
    "Tên hàng hóa/dịch vụ (*)", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền", "Kênh bán hàng",
    "Chiết khấu thương mại", "Hàng khuyến mại",
]
TEMPLATE_WIDTHS = {"A": 18, "B": 8, "C": 14, "D": 22, "E": 16, "F": 24, "G": 22, "H": 16, "I": 18,
                   "K": 65, "L": 47.3, "M": 12, "N": 16, "O": 18, "P": 18.8, "Q": 24, "R": 15.8, "S": 18}
BUYER = "Bán cho người tiêu dùng"
PAYMENT = "TM/CK"
CHECK_HEADERS = ["STT", "Mã GM", "Mã đơn hàng dài", "Ngày tạo", "Số dòng món", "Số dòng chiết khấu",
                 "Tổng SL", "Tổng tạm tính", "Giảm giá thực tế", "Tổng cộng", "Tổng tính lại", "Trạng thái"]
CHECK_WIDTHS = [8, 12, 30, 22, 12, 18, 10, 16, 16, 16, 16, 40]


def invoice_order(results: pd.DataFrame) -> pd.DataFrame:
    """Các đơn ghép được, theo thứ tự xuất hóa đơn (ngày tạo tăng dần), kèm số thứ tự hóa đơn."""
    ok = results[results.ok].sort_values(["created_dt", "order_id"], na_position="last")
    return ok.assign(invoice_no=range(1, len(ok) + 1))


def to_template_excel(results: pd.DataFrame, lines: pd.DataFrame, discount_name: str,
                      invoice_date=None) -> bytes:
    """Xuất theo mẫu nhập phần mềm hóa đơn: dòng món → 1 dòng giảm giá → dòng THÀNH TIỀN.

    invoice_date: None = lấy ngày tạo đơn; hoặc một ngày cố định cho mọi hóa đơn.
    """
    import xlsxwriter

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    base_fmt = {"font_name": "Calibri", "font_size": 11}

    def fmt(**kw):
        return wb.add_format({**base_fmt, **kw})

    title_f = fmt(bold=True, font_size=14)
    plain_f = fmt()
    head_f = fmt(bold=True, align="center", border=1)
    # (định dạng thường, định dạng cột tiền N/O/P) cho từng loại dòng
    styles = {
        "item": (fmt(border=1), fmt(border=1, num_format="#,##0")),
        "discount": (fmt(border=1, bold=True), fmt(border=1, bold=True, num_format="#,##0")),
        "total": (fmt(border=1, bold=True, bg_color="#D9EAF7"),
                  fmt(border=1, bold=True, bg_color="#D9EAF7", num_format="#,##0")),
    }
    money_cols = {13, 14, 15}  # N, O, P (đánh số từ 0)

    ws = wb.add_worksheet("Hóa đơn bán hàng")
    for col, w in TEMPLATE_WIDTHS.items():
        ws.set_column(f"{col}:{col}", w)
    ws.write(0, 1, TEMPLATE_TITLE, title_f)
    for i, text in enumerate(TEMPLATE_GUIDE, start=1):
        ws.write(i, 1, text, plain_f)
    for j, text in enumerate(TEMPLATE_HEADERS, start=1):
        ws.write(7, j, text, head_f)

    row = 8

    def put(values: dict, kind="item"):
        nonlocal row
        text_f, money_f = styles[kind]
        for col in range(19):  # A..S
            f = money_f if col in money_cols else text_f
            v = values.get(col)
            if v is None:
                ws.write_blank(row, col, None, f)
            else:
                ws.write(row, col, v, f)
        row += 1

    order = invoice_order(results)
    by_order: dict[str, list] = {}
    for oid, name, qty, price, total in zip(lines.order_id, lines.name, lines.qty, lines.price, lines.total):
        by_order.setdefault(oid, []).append((name, int(qty), int(price), int(total)))
    checks = []
    for k, o in enumerate(order.itertuples()):
        items = by_order.get(o.order_id, [])
        when = invoice_date if invoice_date is not None else o.created_dt
        date_str = when.strftime("%d/%m/%Y") if pd.notna(when) else ""
        base = {0: o.short_id, 1: o.invoice_no, 2: date_str, 6: BUYER, 10: PAYMENT, 16: f"GRAB - {o.short_id}"}
        if k:
            row += 1  # dòng trống giữa 2 hóa đơn
        for name, qty, price, total in items:
            put({**base, 11: name, 13: qty, 14: price, 15: total})
        discount = -int(round(o.discount))  # giảm giá trong báo cáo là số âm
        if discount != 0:
            put({**base, 11: discount_name, 15: discount, 17: 1}, "discount")
        subtotal = sum(it[3] for it in items)
        net = subtotal - discount
        put({**base, 11: f"THÀNH TIỀN {o.short_id}", 15: net}, "total")

        report_net = None if pd.isna(o.net) else int(round(o.net))
        if subtotal != int(o.amount):
            status = f"LỆCH tổng món {subtotal:,} vs số tiền {int(o.amount):,}"
        elif report_net is None:
            status = "THIẾU DOANH THU RÒNG"
        elif report_net != net:
            status = f"LỆCH {net:,} vs báo cáo {report_net:,}"
        else:
            status = "OK"
        checks.append([o.invoice_no, o.short_id, o.order_id, o.created, len(items), int(discount != 0),
                       sum(it[1] for it in items), int(o.amount), discount, report_net, net, status])
    if row > 8:
        ws.autofilter(7, 0, row - 1, 18)

    for o in results[~results.ok].itertuples():
        checks.append([None, o.short_id, o.order_id, o.created, 0, 0, 0, int(o.amount),
                       -int(round(o.discount)), None if pd.isna(o.net) else int(round(o.net)), None,
                       f"CHƯA GHÉP: {o.fail_reason}"])

    ws2 = wb.add_worksheet("Kiểm tra")
    bold_f, money_f = fmt(bold=True), fmt(num_format="#,##0")
    for j, (h, w) in enumerate(zip(CHECK_HEADERS, CHECK_WIDTHS)):
        ws2.set_column(j, j, w, money_f if 7 <= j <= 10 else plain_f)  # H..K là số tiền
        ws2.write(0, j, h, bold_f)
    for i, r in enumerate(checks, start=1):
        for j, v in enumerate(r):
            if v is not None:
                ws2.write(i, j, v)

    wb.close()
    return buf.getvalue()
