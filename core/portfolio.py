"""損益計算。

匯率採「方案 A」：成本與市值都用**當下同一個匯率**換算成台幣。
因此美股的損益只反映股價漲跌，不含匯差 —— 介面上會明白標註這件事。
"""
from __future__ import annotations

import datetime as _dt

from . import config as C
from .excel_io import Holding
from .providers import FxRate, Quote


def _pct(part: float, whole: float) -> float | None:
    return (part / whole * 100) if whole else None


def build_rows(holdings: list[Holding], quotes: dict[str, Quote],
               fx: FxRate) -> list[dict]:
    """把持股 + 報價組成畫面要的每一列。金額欄位一律已換算成台幣。"""
    rate = fx.rate
    rows: list[dict] = []

    for h in holdings:
        q = quotes.get(h.code) or Quote(code=h.code, error="沒有報價資料")
        is_us = h.market == C.MARKET_US
        conv = (rate if is_us else 1.0)

        price = q.price
        prev = q.prev_close

        row = {
            "代號": h.code,
            "名稱": h.name or q.name or h.code,
            "市場": h.market,
            "股數": h.shares,
            "均價": h.avg_cost,
            "現價": price,
            "昨收": prev,
            "幣別": "USD" if is_us else "TWD",
            "資料來源": q.source,
            "備註": h.note,
            "警告": q.error,
            "遞補價": q.stale,
            # 同一代號有多筆時的序號，畫面用來顯示「①②」避免看起來像重複顯示
            "重複序號": h.dup_index,
            "重複總數": h.dup_total,
        }

        if price is None or (is_us and not conv):
            # 沒報價（或美股缺匯率）就只填得出成本，其餘留空，不要瞎編數字
            row.update({"漲跌幅": None, "今日損益": None, "總損益": None,
                        "報酬率": None, "市值": None,
                        "成本": h.shares * h.avg_cost * (conv or 0) if is_us else h.shares * h.avg_cost})
            if is_us and not conv:
                row["警告"] = (row["警告"] + "；" if row["警告"] else "") + "缺少匯率，無法換算台幣"
                row["成本"] = None
            rows.append(row)
            continue

        market_value = h.shares * price * conv
        cost = h.shares * h.avg_cost * conv
        total_pnl = market_value - cost
        today_pnl = (h.shares * (price - prev) * conv) if prev is not None else None

        row.update({
            "漲跌幅": _pct(price - prev, prev) if prev else None,
            "今日損益": today_pnl,
            "總損益": total_pnl,
            "報酬率": _pct(total_pnl, cost),
            "市值": market_value,
            "成本": cost,
        })
        rows.append(row)

    return rows


def summarize(rows: list[dict], fx: FxRate) -> dict:
    """上方統計卡的數字。全部台幣。"""
    def total(key, pred=lambda r: True):
        vals = [r[key] for r in rows if pred(r) and r.get(key) is not None]
        return sum(vals) if vals else 0.0

    market_value = total("市值")
    cost = total("成本")
    today_pnl = total("今日損益")
    total_pnl = market_value - cost

    tw_value = total("市值", lambda r: r["市場"] == C.MARKET_TW)
    us_value = total("市值", lambda r: r["市場"] == C.MARKET_US)

    incomplete = sorted({r["代號"] for r in rows if r.get("市值") is None})
    no_prev = sorted({r["代號"] for r in rows
                      if r.get("市值") is not None and r.get("今日損益") is None})

    return {
        "market_value": market_value,
        "cost": cost,
        "total_pnl": total_pnl,
        "total_return": _pct(total_pnl, cost),
        "today_pnl": today_pnl,
        "today_return": _pct(today_pnl, market_value - today_pnl),
        "tw_value": tw_value,
        "us_value": us_value,
        "us_ratio": _pct(us_value, market_value),
        "has_us": any(r["市場"] == C.MARKET_US for r in rows),
        "fx_rate": fx.rate,
        "fx_source": fx.source,
        "fx_note": fx.note,
        "fx_error": fx.error,
        "incomplete": incomplete,
        "no_prev_close": no_prev,
        "updated_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def nav_record(summary: dict) -> dict:
    """組成寫進「每日淨值」的一筆。"""
    return {
        "日期": _dt.date.today().strftime("%Y-%m-%d"),
        "總市值": round(summary["market_value"], 2),
        "總成本": round(summary["cost"], 2),
        "總損益": round(summary["total_pnl"], 2),
        "報酬率": round(summary["total_return"], 4) if summary["total_return"] is not None else None,
        "台股市值": round(summary["tw_value"], 2),
        "美股市值(台幣)": round(summary["us_value"], 2),
        "USD/TWD": summary["fx_rate"],
    }
