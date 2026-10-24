"""QC edge-case tests for invoice_engine (spec: docs/quy-tac-ghep-mon.md).

Every test builds small synthetic Excel files in memory and passes the bytes to
load_orders / load_items, like app.py does. The hard rules are checked with an
independent checker (`independent_violations`) that does NOT reuse validate().
"""
from __future__ import annotations

import io
import math
import shutil
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import invoice_engine as eng  # noqa: E402

DATA = ROOT / "data"
STEP_LABELS = [s[0] for s in eng.STEPS]
BRACKETS = [(50_000, 1, 2), (150_000, 2, 4), (300_000, 3, 6), (math.inf, 5, 10)]


# ----------------------------------------------------------------------------- helpers

def xbytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def orders_df(rows):
    """rows: list of dicts with keys id, status, amount, discount, net (net defaults to amount+discount)."""
    out = []
    for i, r in enumerate(rows):
        amount = r.get("amount")
        discount = r.get("discount", 0)
        net = r.get("net", (amount + discount) if isinstance(amount, (int, float)) and amount == amount else None)
        out.append({
            "Tên cửa hàng": "Shop", "Ngày tạo": "2026-01-01", "Loại": "GrabMart",
            "Trạng thái": r.get("status", "Đã _x0008_thanh toán"),
            "Mã đơn hàng dài": r.get("id", f"LONG-{i:04d}"),
            "Mã đơn hàng ngắn": r.get("short", f"GM-{i % 7}"),
            "Mã đặt hàng": f"B{i}", "Phương thức thanh toán": "Cash",
            "Số tiền": amount, "Giảm giá (Quán tài trợ)": discount, "Doanh thu ròng": net,
        })
    return pd.DataFrame(out)


def load_orders_rows(rows):
    return eng.load_orders(xbytes(orders_df(rows)))


def hot_df(items):
    """items: list of (name, price, units_sold)."""
    cols = ["Date", "Item", "Units Sold", "Item Gross Sales (₫)", "Đơn giá"]
    return pd.DataFrame([{"Date": "2026-07-01", "Item": n, "Units Sold": u,
                          "Item Gross Sales (₫)": (p * u if isinstance(p, (int, float)) and p == p else None),
                          "Đơn giá": p} for n, p, u in items], columns=cols)


def cold_df(items):
    """items: list of (name, price, ordered_qty)."""
    cols = ["Ngày", "Sản phẩm", "Số lượng được đặt", "Số lượng thiếu hụt", "Doanh thu bị mất (₫)", "Đơn giá"]
    return pd.DataFrame([{"Ngày": "2026-04-14", "Sản phẩm": n, "Số lượng được đặt": q, "Số lượng thiếu hụt": 1,
                          "Doanh thu bị mất (₫)": (p if isinstance(p, (int, float)) and p == p else None),
                          "Đơn giá": p} for n, p, q in items], columns=cols)


def catalog_from(hot_items, cold_items, exclude=eng.DEFAULT_EXCLUDE):
    return eng.build_catalog(eng.load_items(xbytes(hot_df(hot_items))),
                             eng.load_items(xbytes(cold_df(cold_items))), exclude)


def bounds(amount):
    for limit, lo, hi in BRACKETS:
        if amount < limit:
            return lo, hi


def independent_violations(orders, results, lines, catalog, params, exclude=eng.DEFAULT_EXCLUDE):
    """Own re-implementation of the hard rules (does not call validate())."""
    v = []
    elig = orders[orders.eligible]
    cap = max(1, math.floor(params.cap_pct / 100 * len(elig) + 1e-9))
    cat = {n.casefold(): (int(p), int(t)) for n, p, t in zip(catalog.name, catalog.price, catalog.tier)}
    kws = [k.strip().casefold() for k in exclude if k and k.strip()]

    if sorted(results.order_id) != sorted(elig.order_id):
        v.append("results do not cover each eligible order exactly once")
    res = results.set_index("order_id")
    amount_of = dict(zip(elig.order_id, elig.amount))
    for oid, g in lines.groupby("order_id"):
        if oid not in res.index or not bool(res.loc[oid, "ok"]):
            v.append(f"{oid}: has lines but is not an ok order")
            continue
        recomputed = 0
        for r in g.itertuples():
            if r.name.casefold() not in cat:
                v.append(f"{oid}: unknown item {r.name}")
                continue
            price, tier = cat[r.name.casefold()]
            if r.price != price:
                v.append(f"{oid}: price {r.price} != catalog {price} for {r.name}")
            if r.tier != tier:
                v.append(f"{oid}: tier {r.tier} != catalog {tier} for {r.name}")
            if not (float(r.qty).is_integer() and r.qty >= 1):
                v.append(f"{oid}: bad qty {r.qty}")
            limit = params.max_qty_cheap if price < params.cheap_price else params.max_qty
            if r.qty > limit:
                v.append(f"{oid}: qty {r.qty} > {limit} for {r.name}")
            if r.total != r.price * r.qty:
                v.append(f"{oid}: total != price*qty for {r.name}")
            if any(kw in r.name.casefold() for kw in kws):
                v.append(f"{oid}: excluded item {r.name}")
            if tier == 2 and not str(res.loc[oid, "step"]).startswith("Bước 2"):
                v.append(f"{oid}: tier-2 item used at step {res.loc[oid, 'step']}")
            recomputed += int(price) * int(r.qty)
        if recomputed != int(amount_of[oid]):
            v.append(f"{oid}: lines sum {recomputed} != amount {amount_of[oid]}")
        if g.name.str.casefold().duplicated().any():
            v.append(f"{oid}: item repeated in order")
        if int(res.loc[oid, "n_lines"]) != len(g):
            v.append(f"{oid}: n_lines {res.loc[oid, 'n_lines']} != {len(g)} lines")
    ok_ids = set(results.loc[results.ok == True, "order_id"])  # noqa: E712
    missing = ok_ids - set(lines.order_id)
    if missing:
        v.append(f"ok orders without lines: {sorted(missing)[:5]}")
    per_item = lines.groupby(lines.name.str.casefold()).order_id.nunique()
    for name, n in per_item[per_item > cap].items():
        v.append(f"item {name} in {n} orders > cap {cap}")
    # soft-but-promised rules: line-count bounds and unique combos outside the dup step
    for r in results[results.ok == True].itertuples():  # noqa: E712
        if r.step not in STEP_LABELS:
            v.append(f"{r.order_id}: unknown step {r.step}")
        lo, hi = bounds(r.amount)
        if "nới" in r.step or "trùng" in r.step:
            lo, hi = max(1, lo - 1), hi + max(2, hi // 2)
        if not lo <= r.n_lines <= hi:
            v.append(f"{r.order_id}: {r.n_lines} lines outside {lo}-{hi}")
    sigs = {}
    step_of = dict(zip(results.order_id, results.step))
    for oid, g in lines.groupby("order_id"):
        sigs.setdefault(frozenset(zip(g.name, g.qty)), []).append(oid)
    for sig, oids in sigs.items():
        not_allowed = [o for o in oids if "trùng" not in step_of[o]]
        if len(not_allowed) > 1:
            v.append(f"duplicate combination outside dup step: {not_allowed}")
    return v


def run_all(orders, catalog, **params):
    p = eng.Params(**params)
    res, lines, cap = eng.generate_invoices(orders, catalog, p)
    return p, res, lines, cap


HOT_BASIC = [("Táo", 25_000, 100), ("Cam", 30_000, 80), ("Chuối", 15_000, 60), ("Nho", 55_000, 40),
             ("Dưa hấu", 42_000, 30), ("Xoài", 36_000, 20), ("Rau muống", 8_000, 200), ("Thịt bò", 120_000, 10),
             ("Cá hồi", 189_000, 5), ("Trứng gà", 32_000, 150), ("Sữa tươi", 28_500, 90), ("Bánh mì", 12_000, 70)]
COLD_BASIC = [("Nấm truffle", 7_000, 3), ("Hạt dẻ", 47_000, 2)]


# ============================================================================ order data

class TestOrders:
    def test_cancelled_orders_are_skipped(self):
        o = load_orders_rows([{"id": "A", "status": "Đã hủy", "amount": 50_000},
                              {"id": "B", "amount": 50_000}])
        assert o.set_index("order_id").eligible.to_dict() == {"A": False, "B": True}
        assert "Đã hủy" in o.set_index("order_id").skip_reason["A"]

    def test_unpaid_orders_are_included_other_statuses_skipped(self):
        o = load_orders_rows([{"id": "A", "status": "Chưa thanh toán", "amount": 50_000},
                              {"id": "B", "status": "CHƯA _x0008_thanh toán", "amount": 50_000},
                              {"id": "C", "status": "Thanh toán thất bại", "amount": 50_000},
                              {"id": "D", "status": "", "amount": 50_000}])
        assert o.set_index("order_id").eligible.to_dict() == {"A": True, "B": True, "C": False, "D": False}

    def test_status_with_literal_x0008_artifact_is_eligible(self):
        o = load_orders_rows([{"id": "A", "status": "Đã _x0008_thanh toán", "amount": 50_000}])
        assert o.eligible.all()
        assert o.status.iat[0] == "Đã thanh toán"

    def test_status_with_real_backspace_char_is_eligible(self, monkeypatch):
        # openpyxl refuses to write \x08, so feed the frame straight into the reader
        df = orders_df([{"id": "A", "status": "Đã \x08thanh toán", "amount": 50_000}])
        monkeypatch.setattr(eng.pd, "read_excel", lambda *a, **k: df.copy())
        o = eng.load_orders(b"ignored")
        assert o.eligible.all() and o.status.iat[0] == "Đã thanh toán"

    @pytest.mark.parametrize("amount", [None, float("nan"), 0, -15_000])
    def test_blank_zero_negative_amount_skipped(self, amount):
        o = load_orders_rows([{"id": "A", "amount": amount}])
        assert not o.eligible.iat[0]
        assert "Số tiền" in o.skip_reason.iat[0]

    @pytest.mark.parametrize("amount", [100_000.5, 32_537.5, 99_999.01])
    def test_fractional_amount_skipped(self, amount):
        o = load_orders_rows([{"id": "A", "amount": amount}])
        assert not o.eligible.iat[0] and o.skip_reason.iat[0] == "Số tiền có phần lẻ"

    def test_tiny_positive_fractional_amount_not_eligible_with_zero(self):
        # 1e-7 is a fractional amount; it must be skipped, never become an eligible order of 0đ
        o = load_orders_rows([{"id": "A", "amount": 1e-7}])
        assert not (o.eligible.iat[0] and o.amount.iat[0] <= 0), o[["amount", "eligible", "skip_reason"]].to_dict("records")

    def test_duplicate_long_order_id_second_skipped(self):
        o = load_orders_rows([{"id": "DUP", "amount": 50_000}, {"id": "DUP", "amount": 70_000},
                              {"id": " DUP\xa0", "amount": 90_000}])
        assert o.eligible.tolist() == [True, False, False]
        assert o.skip_reason.iat[1].startswith("Trùng mã đơn với dòng 2")

    def test_duplicate_short_id_is_fine(self):
        o = load_orders_rows([{"id": "A", "short": "GM-1", "amount": 50_000},
                              {"id": "B", "short": "GM-1", "amount": 50_000}])
        assert o.eligible.all()

    def test_missing_long_id_skipped(self):
        o = load_orders_rows([{"id": None, "amount": 50_000}, {"id": "   ", "amount": 50_000}])
        assert not o.eligible.any() and (o.skip_reason == "Thiếu mã đơn hàng dài").all()

    def test_only_required_columns(self):
        df = pd.DataFrame({"Mã đơn hàng dài": ["A"], "Trạng thái": ["Đã thanh toán"], "Số tiền": [50_000]})
        o = eng.load_orders(xbytes(df))
        assert o.eligible.iat[0] and o.discount.iat[0] == 0
        assert "Thiếu doanh thu ròng" in o.warning.iat[0]

    @pytest.mark.parametrize("drop", [["Số tiền"], ["Trạng thái"], ["Mã đơn hàng dài"], ["Số tiền", "Trạng thái"]])
    def test_missing_required_columns_raise_value_error(self, drop):
        df = orders_df([{"id": "A", "amount": 50_000}]).drop(columns=drop)
        with pytest.raises(ValueError) as e:
            eng.load_orders(xbytes(df))
        for c in drop:
            assert c in str(e.value)

    def test_header_with_line_break_is_recognised(self):
        # Excel headers often contain Alt+Enter line breaks; "Số\ntiền" should read as "Số tiền"
        df = pd.DataFrame({"Mã đơn hàng dài": ["A"], "Trạng thái": ["Đã thanh toán"], "Số\ntiền": [50_000]})
        o = eng.load_orders(xbytes(df))
        assert o.eligible.iat[0]

    def test_net_mismatch_warns_but_eligible(self):
        o = load_orders_rows([{"id": "A", "amount": 100_000, "discount": -10_000, "net": 95_000}])
        assert o.eligible.iat[0] and "Doanh thu ròng ≠ Số tiền + Giảm giá" in o.warning.iat[0]

    def test_positive_discount_warns(self):
        o = load_orders_rows([{"id": "A", "amount": 100_000, "discount": 10_000}])
        assert o.eligible.iat[0] and "Giảm giá là số dương" in o.warning.iat[0]

    def test_discount_bigger_than_amount_warns(self):
        o = load_orders_rows([{"id": "A", "amount": 100_000, "discount": -150_000}])
        assert o.eligible.iat[0] and "Giảm giá lớn hơn số tiền" in o.warning.iat[0]

    def test_clean_order_has_no_warning(self):
        o = load_orders_rows([{"id": "A", "amount": 100_000, "discount": -10_000}])
        assert o.warning.iat[0] == ""

    def test_amount_float_noise_is_rounded(self):
        o = load_orders_rows([{"id": "A", "amount": 100_000.0000001}])
        assert o.eligible.iat[0] and o.amount.iat[0] == 100_000


# ============================================================================ item data

class TestItems:
    def test_vietnamese_header_row_inside_data_is_dropped(self):
        hot = pd.concat([pd.DataFrame([{"Date": "Ngày", "Item": "Mục", "Units Sold": "Số lượng sản phẩm đã bán",
                                        "Item Gross Sales (₫)": "Doanh thu gộp của mặt hàng (₫)", "Đơn giá": "Đơn giá"}]),
                         hot_df([("Táo", 25_000, 3)])], ignore_index=True)
        items = eng.load_items(xbytes(hot))
        assert items.name.tolist() == ["Táo"]

    def test_name_variants_merge_into_one_item(self):
        hot = hot_df([('"Cà chua bi', 24_000, 5), ("Cà\xa0chua  bi ", 24_000, 3), ("CÀ CHUA BI", 24_000, 1),
                      ('  "cà chua bi', 24_000, 1)])
        cat, _ = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold_df([]))))
        assert len(cat) == 1
        assert cat.weight.iat[0] == 10
        assert '"' not in cat.name.iat[0]

    def test_name_with_tab_or_newline_matches_spaced_name(self):
        # control whitespace (tab/newline) must behave like a space, not glue words together
        assert eng.clean_text("Bánh\nmì") == "Bánh mì"
        assert eng.clean_text("Coca\tCola") == "Coca Cola"

    def test_price_mode_by_units_sold(self):
        hot = hot_df([("Táo", 25_000, 1), ("Táo", 25_000, 1), ("Táo", 25_000, 1), ("Táo", 27_000, 10)])
        cat, _ = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold_df([]))))
        assert cat.price.iat[0] == 27_000 and cat.n_prices.iat[0] == 2

    def test_price_tie_is_deterministic(self):
        hot = hot_df([("Táo", 25_000, 5), ("Táo", 27_000, 5)])
        a, _ = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold_df([]))))
        b, _ = eng.build_catalog(eng.load_items(xbytes(hot.iloc[::-1])), eng.load_items(xbytes(cold_df([]))))
        assert a.price.iat[0] in (25_000, 27_000)
        assert a.price.iat[0] == b.price.iat[0]

    def test_non_integer_prices_ignored(self):
        hot = hot_df([("Táo", 32_537.5, 100), ("Táo", 30_000, 1), ("Lê", 12_345.6, 3)])
        cat, excl = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold_df([]))))
        assert cat.set_index("name").price.to_dict() == {"Táo": 30_000}
        assert "Lê" in excl["Món"].tolist()

    def test_blank_price_falls_back_to_revenue_over_qty(self):
        hot = pd.DataFrame([{"Item": "Táo", "Units Sold": 4, "Item Gross Sales (₫)": 100_000, "Đơn giá": None},
                            {"Item": "Lê", "Units Sold": 3, "Item Gross Sales (₫)": 100_000, "Đơn giá": None},
                            {"Item": "Mận", "Units Sold": 0, "Item Gross Sales (₫)": 100_000, "Đơn giá": None}])
        cat, excl = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold_df([]))))
        assert cat.set_index("name").price.to_dict() == {"Táo": 25_000}
        assert "Lê" in excl["Món"].tolist()  # 33,333.33 is fractional

    def test_blank_price_column_missing_entirely_uses_fallback(self):
        hot = pd.DataFrame([{"Item": "Táo", "Units Sold": 4, "Item Gross Sales (₫)": 100_000}])
        items = eng.load_items(xbytes(hot))
        assert items.price.tolist() == [25_000]

    @pytest.mark.parametrize("bad", [0, -5_000])
    def test_zero_negative_price_ignored(self, bad):
        hot = hot_df([("Táo", bad, 100), ("Táo", 20_000, 1), ("Lê", bad, 1)])
        cat, excl = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold_df([]))))
        assert cat.set_index("name").price.to_dict() == {"Táo": 20_000}
        assert "Lê" in excl["Món"].tolist()

    def test_tier1_price_wins_over_tier2(self):
        cat, _ = catalog_from([("Táo", 25_000, 5)], [("táo ", 99_000, 50), ("Hạt dẻ", 47_000, 2)])
        c = cat.set_index("key")
        assert c.loc["táo", "price"] == 25_000 and c.loc["táo", "tier"] == 1
        assert c.loc["hạt dẻ", "tier"] == 2
        assert len(cat) == 2

    def test_tier2_weight_is_ordered_qty_min_1(self):
        cat, _ = catalog_from([("Táo", 25_000, 0)], [("Hạt dẻ", 47_000, 0), ("Nấm", 7_000, 6)])
        w = cat.set_index("name").weight.to_dict()
        assert w == {"Táo": 1, "Hạt dẻ": 1, "Nấm": 6}

    def test_default_exclude_gift_items(self):
        cat, excl = catalog_from([("[Hàng tặng không bán] Khay trái cây", 50_000, 9), ("Táo", 25_000, 1)],
                                 [("[HÀNG TẶNG không bán] Quà", 10_000, 1)])
        assert cat.name.tolist() == ["Táo"]
        assert len(excl) == 2

    def test_custom_exclude_keywords_case_insensitive_and_blank_ignored(self):
        cat, _ = catalog_from([("Táo FLASH SALE", 25_000, 1), ("Cam", 30_000, 1)], [], ["flash sale", "", "  "])
        assert cat.name.tolist() == ["Cam"]

    def test_exclude_matching_everything_gives_value_error_not_crash(self):
        cat, _ = catalog_from([("Táo", 25_000, 1)], [("Cam", 30_000, 1)], ["táo", "CAM"])
        assert cat.empty
        orders = load_orders_rows([{"id": "A", "amount": 50_000}])
        with pytest.raises(ValueError):
            eng.generate_invoices(orders, cat, eng.Params())

    def test_item_files_without_any_valid_price_give_value_error(self):
        # both files have no usable rows -> should be a clear ValueError (app only catches ValueError)
        hot = hot_df([("Táo", 32_537.5, 1)])
        cold = cold_df([("Cam", -1, 1)])
        with pytest.raises(ValueError):
            cat, _ = eng.build_catalog(eng.load_items(xbytes(hot)), eng.load_items(xbytes(cold)))
            eng.generate_invoices(load_orders_rows([{"id": "A", "amount": 50_000}]), cat, eng.Params())

    def test_exclude_keyword_is_normalised_like_names(self):
        # names are NFC-normalised and whitespace-collapsed; keywords typed by the user should be too
        import unicodedata
        for kw in (unicodedata.normalize("NFD", "Hàng tặng"), "Hàng  tặng"):
            cat, _ = catalog_from([("[Hàng tặng không bán] Khay", 50_000, 1), ("Táo", 25_000, 1)], [], [kw])
            assert cat.name.tolist() == ["Táo"], repr(kw)

    def test_vietnamese_header_item_files(self):
        hot = pd.DataFrame({"Mục": ["Táo"], "Số lượng sản phẩm đã bán": [4],
                            "Doanh thu gộp của mặt hàng (₫)": [100_000], "Đơn giá": [25_000]})
        items = eng.load_items(xbytes(hot))
        assert items.to_dict("records") == [{"name": "Táo", "price": 25_000, "weight": 4}]

    def test_item_not_reported_dropped_when_still_in_catalog(self):
        # hot row has only a fractional price, cold row has a valid one -> item is usable (tier 2),
        # so the "Món bị loại" list must not also claim it was dropped
        cat, excl = catalog_from([("Táo", 32_537.5, 10)], [("Táo", 30_000, 1)])
        assert cat.set_index("name").loc["Táo", "tier"] == 2
        assert "Táo" not in excl["Món"].tolist()

    def test_item_file_missing_columns_raises(self):
        with pytest.raises(ValueError):
            eng.load_items(xbytes(pd.DataFrame({"Foo": ["x"], "Đơn giá": [1000]})))
        with pytest.raises(ValueError):
            eng.load_items(xbytes(pd.DataFrame({"Item": ["x"], "Units Sold": [1]})))


# ============================================================================ generation

class TestGeneration:
    def test_amount_below_cheapest_item_fails_with_reason(self):
        cat, _ = catalog_from(HOT_BASIC, COLD_BASIC)
        orders = load_orders_rows([{"id": "A", "amount": 5_000}])
        p, res, lines, cap = run_all(orders, cat)
        assert not res.ok.iat[0] and "nhỏ hơn" in res.fail_reason.iat[0]
        assert lines.empty

    def test_tier2_only_amount_succeeds_at_step2_never_earlier(self):
        cat, _ = catalog_from([("A", 10_000, 50), ("B", 20_000, 50)], [("Z", 7_000, 1)])
        orders = load_orders_rows([{"id": "O1", "amount": 7_000}, {"id": "O2", "amount": 17_000},
                                   {"id": "O3", "amount": 30_000}])
        p, res, lines, cap = run_all(orders, cat, cap_pct=100)
        r = res.set_index("order_id")
        assert r.ok.all(), r[["ok", "fail_reason"]]
        for oid in ("O1", "O2"):
            assert r.loc[oid, "step"].startswith("Bước 2")
            assert (lines[lines.order_id == oid].tier == 2).any()
        assert r.loc["O3", "step"].startswith("Bước 1")
        assert independent_violations(orders, res, lines, cat, p) == []

    @pytest.mark.parametrize("amount", [15_000, 25_000, 1_000_000])
    def test_impossible_amount_fails_cleanly(self, amount):
        # only multiples of 10,000 are reachable, and 1,000,000 needs > 15 lines x 3 x 20k
        cat, _ = catalog_from([("A", 10_000, 50), ("B", 20_000, 50)], [])
        orders = load_orders_rows([{"id": "X", "amount": amount}])
        p, res, lines, cap = run_all(orders, cat, cap_pct=100, max_qty_cheap=3)
        assert not res.ok.iat[0]
        assert res.fail_reason.iat[0] == "Không tìm được tổ hợp món khớp đúng số tiền"
        assert lines.empty

    def test_huge_amount_fails_fast(self):
        cat, _ = catalog_from(HOT_BASIC, COLD_BASIC)
        orders = load_orders_rows([{"id": "X", "amount": 987_654_000}])
        t = time.time()
        p, res, lines, cap = run_all(orders, cat)
        assert time.time() - t < 30
        assert not res.ok.iat[0] and lines.empty

    def test_cheap_qty_limit_boundary(self):
        # price < 20k -> max 5; price == 20k -> max 3
        cat, _ = catalog_from([("Rẻ", 1_000, 1)], [])
        orders = load_orders_rows([{"id": "A", "amount": 5_000}, {"id": "B", "amount": 6_000}])
        p, res, lines, cap = run_all(orders, cat, cap_pct=100)
        r = res.set_index("order_id")
        assert r.loc["A", "ok"] and lines.qty.max() == 5
        assert not r.loc["B", "ok"]

        cat2, _ = catalog_from([("Vừa", 20_000, 1)], [])
        orders2 = load_orders_rows([{"id": "A", "amount": 60_000}, {"id": "B", "amount": 80_000}])
        p, res2, lines2, _ = run_all(orders2, cat2, cap_pct=100)
        r2 = res2.set_index("order_id")
        assert r2.loc["A", "ok"] and not r2.loc["B", "ok"]

    def test_cap_with_three_orders_is_one(self):
        items = [(f"M{i}", 10_000 + 1_000 * i, 10) for i in range(15)]
        cat, _ = catalog_from(items, [])
        orders = load_orders_rows([{"id": f"O{i}", "amount": a} for i, a in enumerate([31_000, 45_000, 60_000])])
        p, res, lines, cap = run_all(orders, cat, cap_pct=10)
        assert cap == 1
        assert lines.groupby("name").order_id.nunique().max() <= 1
        assert independent_violations(orders, res, lines, cat, p) == []

    def test_forced_duplicate_combination(self):
        cat, _ = catalog_from([("Duy nhất", 10_000, 1)], [])
        orders = load_orders_rows([{"id": f"O{i}", "amount": 10_000} for i in range(3)])
        p, res, lines, cap = run_all(orders, cat, cap_pct=100)
        r = res.set_index("order_id")
        assert r.ok.all()
        assert r.loc["O0", "step"] == "Bước 1"
        assert r.loc["O1", "step"] == r.loc["O2", "step"] == "Bước 2 (cho phép trùng tổ hợp)"
        assert eng.validate(res, lines, cat, p, cap) == []
        assert independent_violations(orders, res, lines, cat, p) == []

    def test_forced_duplicate_blocked_by_cap(self):
        cat, _ = catalog_from([("Duy nhất", 10_000, 1)], [])
        orders = load_orders_rows([{"id": f"O{i}", "amount": 10_000} for i in range(3)])
        p, res, lines, cap = run_all(orders, cat, cap_pct=10)
        assert res.ok.sum() == 1
        assert lines.groupby("name").order_id.nunique().max() == 1
        assert (res[~res.ok].fail_reason != "").all()

    def test_zero_eligible_orders(self):
        cat, _ = catalog_from(HOT_BASIC, COLD_BASIC)
        orders = load_orders_rows([{"id": "A", "status": "Đã hủy", "amount": 50_000}])
        p, res, lines, cap = run_all(orders, cat)
        assert res.empty and lines.empty
        assert eng.validate(res, lines, cat, p, cap) == []

    def test_all_orders_fail_validate_does_not_crash(self):
        cat, _ = catalog_from(HOT_BASIC, COLD_BASIC)
        orders = load_orders_rows([{"id": "A", "amount": 1_000}, {"id": "B", "amount": 2_000}])
        p, res, lines, cap = run_all(orders, cat)
        assert not res.ok.any() and lines.empty
        assert eng.validate(res, lines, cat, p, cap) == []
        eng.to_excel_bytes(res, lines, eng.item_usage(lines, 0))

    def test_determinism_same_seed(self):
        cat, _ = catalog_from(HOT_BASIC, COLD_BASIC)
        rng = np.random.default_rng(0)
        orders = load_orders_rows([{"id": f"O{i:03d}", "amount": int(rng.integers(3, 80)) * 5_000} for i in range(60)])
        a = run_all(orders, cat, seed=11)
        b = run_all(orders, cat, seed=11)
        pd.testing.assert_frame_equal(a[1], b[1])
        pd.testing.assert_frame_equal(a[2], b[2])
        c = run_all(orders, cat, seed=12)
        assert not a[2].equals(c[2])

    @pytest.mark.parametrize("seed,cap_pct", [(1, 10), (2, 3), (3, 1), (4, 50)])
    def test_random_property_hard_rules(self, seed, cap_pct):
        rng = np.random.default_rng(seed)
        hot = [(f"Hot {i}", int(rng.integers(5, 200)) * 1_000 + int(rng.choice([0, 500])), int(rng.integers(0, 300)))
               for i in range(60)]
        cold = [(f"Cold {i}", int(rng.integers(5, 120)) * 1_000 + 300, int(rng.integers(0, 5))) for i in range(15)]
        cat, _ = catalog_from(hot, cold)
        prices = cat.price.to_numpy()
        amounts = []
        for _ in range(300):
            k = int(rng.integers(1, 8))
            amounts.append(int(sum(prices[rng.choice(len(prices), k, replace=False)] * rng.integers(1, 3, k))))
        amounts += [int(rng.integers(1, 400)) * 1_000 + 7 for _ in range(20)]  # mostly impossible
        orders = load_orders_rows([{"id": f"O{i:04d}", "amount": a} for i, a in enumerate(amounts)])
        p, res, lines, cap = run_all(orders, cat, seed=seed, cap_pct=cap_pct)
        assert independent_violations(orders, res, lines, cat, p) == []
        assert eng.validate(res, lines, cat, p, cap) == []
        assert (res[~res.ok].fail_reason != "").all()


# ============================================================================ the checker itself

class TestValidateCatchesViolations:
    @pytest.fixture(scope="class")
    @classmethod
    def good(cls):
        cat, _ = catalog_from(HOT_BASIC, COLD_BASIC)
        rng = np.random.default_rng(5)
        orders = load_orders_rows([{"id": f"O{i:03d}", "amount": int(rng.integers(4, 60)) * 5_000} for i in range(40)])
        p, res, lines, cap = run_all(orders, cat, cap_pct=50)
        assert res.ok.sum() >= 30
        assert eng.validate(res, lines, cat, p, cap) == []
        return cat, p, res, lines, cap

    def _first_ok(self, res, lines, min_lines=1):
        counts = lines.groupby("order_id").size()
        return next(o for o in res[res.ok].order_id if counts.get(o, 0) >= min_lines)

    def test_catches_wrong_sum(self, good):
        cat, p, res, lines, cap = good
        bad = lines.copy()
        i = bad.index[0]
        bad.loc[i, "total"] += 1
        assert any("≠ số tiền" in s for s in eng.validate(res, bad, cat, p, cap))

    def test_catches_wrong_price(self, good):
        cat, p, res, lines, cap = good
        bad = lines.copy()
        bad.loc[bad.index[0], "price"] += 1000
        assert any("không khớp danh mục" in s for s in eng.validate(res, bad, cat, p, cap))

    def test_catches_qty_over_limit_and_below_1(self, good):
        cat, p, res, lines, cap = good
        bad = lines.copy()
        bad.loc[bad.index[0], "qty"] = 9
        bad.loc[bad.index[1], "qty"] = 0
        out = eng.validate(res, bad, cat, p, cap)
        assert any("vượt mức tối đa" in s for s in out) and any("< 1" in s for s in out)

    def test_catches_repeated_item_in_order(self, good):
        cat, p, res, lines, cap = good
        bad = pd.concat([lines, lines.iloc[[0]]], ignore_index=True)
        assert any("bị lặp" in s for s in eng.validate(res, bad, cat, p, cap))

    def test_catches_cap_exceeded(self, good):
        cat, p, res, lines, cap = good
        assert any("vượt trần" in s for s in eng.validate(res, lines, cat, p, cap=1))

    def test_catches_excluded_item(self, good):
        cat, p, res, lines, cap = good
        name = lines.name.iat[0]
        out = eng.validate(res, lines, cat, p, cap, exclude_keywords=[name[:3]])
        assert any("loại trừ" in s for s in out)

    def test_catches_duplicate_signature(self, good):
        cat, p, res, lines, cap = good
        a, b = res[res.ok].order_id.iloc[:2]
        copy_a = lines[lines.order_id == a].assign(order_id=b)
        bad = pd.concat([lines[lines.order_id != b], copy_a], ignore_index=True)
        res2 = res.copy()
        res2.loc[res2.order_id == b, "amount"] = res2.loc[res2.order_id == a, "amount"].iat[0]
        res2.loc[res2.order_id == b, "n_lines"] = len(copy_a)
        res2.loc[res2.order_id == b, "step"] = "Bước 1"
        assert any("trùng tổ hợp" in s for s in eng.validate(res2, bad, cat, p, cap))

    def test_catches_line_count_out_of_bounds(self, good):
        cat, p, res, lines, cap = good
        res2 = res.copy()
        res2.loc[res2.ok, "n_lines"] = 99
        # validate đếm số dòng thực tế; n_lines bị sửa thì phải báo lệch ghi nhận
        assert any("thực tế có" in s for s in eng.validate(res2, lines, cat, p, cap))
        # số dòng thực tế vượt khoảng: nhân bản một dòng món sang tên khác trong cùng đơn
        oid = res.loc[res.ok].order_id.iat[0]
        extra = pd.concat([lines[lines.order_id == oid]] * 12, ignore_index=True)
        extra["name"] = [f"X{i}" for i in range(len(extra))]
        bad = pd.concat([lines[lines.order_id != oid], extra], ignore_index=True)
        assert any("ngoài khoảng" in s for s in eng.validate(res, bad, cat, p, cap))

    def test_catches_total_not_equal_price_times_qty(self, good):
        # line shows qty 2 x price but its 'total' still counts 1 -> invoice arithmetic is wrong (rule 1/2)
        cat, p, res, lines, cap = good
        bad = lines.copy()
        i = bad.index[(bad.qty == 1) & (bad.price >= p.cheap_price)][0]
        bad.loc[i, "qty"] = 2
        assert eng.validate(res, bad, cat, p, cap) != []

    def test_catches_non_integer_qty(self, good):
        cat, p, res, lines, cap = good
        bad = lines.copy().astype({"qty": float, "total": float})
        oid = self._first_ok(res, lines, min_lines=2)
        idx = bad.index[bad.order_id == oid]
        i, j = idx[0], idx[1]
        # move half a unit of value between two lines while keeping the order total intact
        bad.loc[i, "qty"] += 0.5
        bad.loc[i, "total"] = bad.loc[i, "price"] * bad.loc[i, "qty"]
        bad.loc[j, "total"] -= bad.loc[i, "price"] * 0.5
        assert eng.validate(res, bad, cat, p, cap) != []

    def test_catches_lines_attached_to_failed_or_unknown_order(self, good):
        cat, p, res, lines, cap = good
        extra = lines.iloc[[0]].assign(order_id="GHOST")
        bad = pd.concat([lines, extra], ignore_index=True)
        assert eng.validate(res, bad, cat, p, cap) != []


# ============================================================================ real data

REAL_FILES = [DATA / "BAO CAO T1 - XULY.xlsx", DATA / "HÀNG BÁN CHẠY.xlsx", DATA / "HÀNG BÁN KHÔNG CHẠY.xlsx"]


@pytest.fixture(scope="module")
def real():
    if not all(f.exists() for f in REAL_FILES):
        pytest.skip("real data not present")
    orders = eng.load_orders(REAL_FILES[0].read_bytes())
    cat, excl = eng.build_catalog(eng.load_items(REAL_FILES[1].read_bytes()), eng.load_items(REAL_FILES[2].read_bytes()))
    return orders, cat, excl


@pytest.mark.parametrize("seed,cap_pct", [(2026, 10), (1, 1), (7, 3), (99, 1)])
def test_real_data_hard_rules(real, seed, cap_pct):
    orders, cat, _ = real
    p, res, lines, cap = run_all(orders, cat, seed=seed, cap_pct=cap_pct)
    assert cap == max(1, math.floor(cap_pct / 100 * orders.eligible.sum()))
    assert independent_violations(orders, res, lines, cat, p) == []
    assert eng.validate(res, lines, cat, p, cap) == []
    assert (res[~res.ok].fail_reason != "").all()


def test_real_data_load_summary(real):
    orders, cat, excl = real
    assert orders.eligible.sum() + (~orders.eligible).sum() == len(orders)
    assert set(orders.loc[orders.eligible, "status"]) == {"Đã thanh toán"}
    assert not cat.name.str.casefold().str.contains("hàng tặng").any()
    assert cat.key.is_unique and (cat.price > 0).all()


# ============================================================================ app smoke tests

def _app_in(tmp_path, orders: pd.DataFrame | None = None, hot=None, cold=None):
    shutil.copy(ROOT / "app.py", tmp_path / "app.py")
    shutil.copy(ROOT / "invoice_engine.py", tmp_path / "invoice_engine.py")
    d = tmp_path / "data"
    d.mkdir()
    if orders is None:
        for f in REAL_FILES:
            shutil.copy(f, d / f.name)
    else:
        (d / "BAO CAO T1 - XULY.xlsx").write_bytes(xbytes(orders))
        (d / "HÀNG BÁN CHẠY.xlsx").write_bytes(xbytes(hot))
        (d / "HÀNG BÁN KHÔNG CHẠY.xlsx").write_bytes(xbytes(cold))
    from streamlit.testing.v1 import AppTest
    return AppTest.from_file(str(tmp_path / "app.py"), default_timeout=300)


def test_app_all_orders_cancelled(tmp_path):
    at = _app_in(tmp_path, orders_df([{"id": "A", "status": "Đã hủy", "amount": 50_000}]),
                 hot_df(HOT_BASIC), cold_df(COLD_BASIC))
    at.run()
    assert not at.exception, at.exception


def test_app_exclude_everything_shows_error(tmp_path):
    at = _app_in(tmp_path, orders_df([{"id": "A", "amount": 50_000}]), hot_df(HOT_BASIC), cold_df(COLD_BASIC))
    at.run()
    assert not at.exception
    at.sidebar.text_input[0].set_value("a, e, o, u, i, n").run()
    assert not at.exception, at.exception
    assert at.error, "expected a st.error message"


def test_app_item_files_without_valid_prices_shows_error(tmp_path):
    at = _app_in(tmp_path, orders_df([{"id": "A", "amount": 50_000}]),
                 hot_df([("Táo", 32_537.5, 1)]), cold_df([("Cam", -1, 1)]))
    at.run()
    assert not at.exception, at.exception
    assert at.error


def test_app_real_data_smoke(tmp_path):
    if not all(f.exists() for f in REAL_FILES):
        pytest.skip("real data not present")
    at = _app_in(tmp_path)
    at.run()
    assert not at.exception, at.exception
    assert not at.error, [e.value for e in at.error]
    assert at.success  # validate() reported 0 violations
