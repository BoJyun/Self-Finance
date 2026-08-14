"""損益計算的單元測試。算錯錢最傷，所以這層要有測試。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core import config as C
from core.excel_io import Holding, infer_market
from core.portfolio import build_rows, summarize
from core.providers import FxRate, Quote

PASS = FAIL = 0


def check(name, got, want, tol=1e-6):
    global PASS, FAIL
    if want is None:
        ok = got is None
    elif got is None:
        ok = False
    else:
        ok = abs(got - want) <= tol
    if ok:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        print(f"  FAIL  {name}: got {got!r}, want {want!r}")


def check_true(name, cond):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        print(f"  FAIL  {name}")


print("── 台股基本損益 ──")
h = [Holding(code="2882", name="國泰金", market=C.MARKET_TW, shares=217, avg_cost=44.66)]
q = {"2882": Quote(code="2882", price=101.5, prev_close=99.5, currency="TWD")}
rows = build_rows(h, q, FxRate(rate=32.0))
r = rows[0]
check("市值 = 217 x 101.5", r["市值"], 22025.5)
check("成本 = 217 x 44.66", r["成本"], 9691.22)
check("總損益", r["總損益"], 22025.5 - 9691.22)
check("報酬率 %", r["報酬率"], (22025.5 - 9691.22) / 9691.22 * 100)
check("今日損益 = 217 x (101.5-99.5)", r["今日損益"], 434.0)
check("漲跌幅 %", r["漲跌幅"], (101.5 - 99.5) / 99.5 * 100)

print("\n── 美股：方案 A，成本與市值用同一個匯率 ──")
h = [Holding(code="AAPL", name="Apple", market=C.MARKET_US, shares=50, avg_cost=180.5)]
q = {"AAPL": Quote(code="AAPL", price=305.26, prev_close=302.25, currency="USD")}
rows = build_rows(h, q, FxRate(rate=32.0))
r = rows[0]
check("市值台幣 = 50 x 305.26 x 32", r["市值"], 50 * 305.26 * 32)
check("成本台幣 = 50 x 180.5 x 32", r["成本"], 50 * 180.5 * 32)
check("今日損益台幣", r["今日損益"], 50 * (305.26 - 302.25) * 32)
check("報酬率不受匯率影響（方案 A 的定義）",
      r["報酬率"], (305.26 - 180.5) / 180.5 * 100)

print("\n  匯率換一個數字，報酬率應完全不變：")
r2 = build_rows(h, q, FxRate(rate=28.0))[0]
check("報酬率 @32 == 報酬率 @28", r2["報酬率"], r["報酬率"])
check_true("市值 @28 < 市值 @32", r2["市值"] < r["市值"])

print("\n── 缺報價時不要瞎編數字 ──")
h = [Holding(code="9999", market=C.MARKET_TW, shares=100, avg_cost=10, name="")]
rows = build_rows(h, {"9999": Quote(code="9999", error="查無此代號")}, FxRate(rate=32.0))
r = rows[0]
check("市值留空", r["市值"], None)
check("總損益留空", r["總損益"], None)
check("成本仍算得出來", r["成本"], 1000.0)
check_true("有帶出錯誤訊息", "查無" in r["警告"])

print("\n── 美股缺匯率 ──")
h = [Holding(code="AAPL", market=C.MARKET_US, shares=10, avg_cost=100, name="")]
rows = build_rows(h, {"AAPL": Quote(code="AAPL", price=200, prev_close=190)},
                  FxRate(rate=None, error="抓不到"))
r = rows[0]
check("市值留空", r["市值"], None)
check("成本也留空（沒匯率就換算不出台幣）", r["成本"], None)
check_true("警告提到匯率", "匯率" in r["警告"])

print("\n── 缺昨收：總損益要算得出來，今日損益留空 ──")
h = [Holding(code="0050", market=C.MARKET_TW, shares=100, avg_cost=100, name="")]
rows = build_rows(h, {"0050": Quote(code="0050", price=120, prev_close=None)}, FxRate(rate=32.0))
r = rows[0]
check("總損益 = 2000", r["總損益"], 2000.0)
check("今日損益留空", r["今日損益"], None)
check("漲跌幅留空", r["漲跌幅"], None)

print("\n── 統計卡 ──")
h = [
    Holding(code="2882", market=C.MARKET_TW, shares=217, avg_cost=44.66, name=""),
    Holding(code="AAPL", market=C.MARKET_US, shares=50, avg_cost=180.5, name=""),
    Holding(code="9999", market=C.MARKET_TW, shares=100, avg_cost=10, name=""),   # 無報價
]
q = {
    "2882": Quote(code="2882", price=101.5, prev_close=99.5),
    "AAPL": Quote(code="AAPL", price=305.26, prev_close=302.25, currency="USD"),
    "9999": Quote(code="9999", error="查無"),
}
rows = build_rows(h, q, FxRate(rate=32.0))
s = summarize(rows, FxRate(rate=32.0))
tw_v = 217 * 101.5
us_v = 50 * 305.26 * 32
check("總市值 = 台股 + 美股（無報價那檔不計入）", s["market_value"], tw_v + us_v)
check("台股市值", s["tw_value"], tw_v)
check("美股市值(台幣)", s["us_value"], us_v)
check("今日損益合計", s["today_pnl"], 434.0 + 50 * (305.26 - 302.25) * 32)
check("總損益 = 市值 - 成本", s["total_pnl"], s["market_value"] - s["cost"])
check_true("有標出算不完整的標的", s["incomplete"] == ["9999"])
check_true("偵測到持有美股", s["has_us"] is True)
check("美股佔比 %", s["us_ratio"], us_v / (tw_v + us_v) * 100)

print("\n── 市場自動判斷 ──")
for code, want in [("2882", C.MARKET_TW), ("00919", C.MARKET_TW), ("006208", C.MARKET_TW),
                   ("AAPL", C.MARKET_US), ("VOO", C.MARKET_US), ("BRK.B", C.MARKET_US)]:
    check_true(f"{code} -> {want}", infer_market(code) == want)

print("\n── 截圖那 5 檔的實際試算（用 2026-08-14 收盤價）──")
screenshot = [
    ("2882", 217, 44.66, 101.5, 99.5),
    ("00919", 38000, 22.24, 30.57, 30.41),
    ("00878", 46000, 23.90, 33.87, 33.78),
    ("006208", 308, 91.48, 243.45, 244.40),
    ("2885", 3518, 27.10, 69.70, 69.70),
]
h = [Holding(code=c, market=C.MARKET_TW, shares=sh, avg_cost=ac, name="") for c, sh, ac, _, _ in screenshot]
q = {c: Quote(code=c, price=p, prev_close=y) for c, _, _, p, y in screenshot}
rows = build_rows(h, q, FxRate(rate=32.0))
s = summarize(rows, FxRate(rate=32.0))
want_mv = sum(sh * p for _, sh, _, p, _ in screenshot)
want_today = sum(sh * (p - y) for _, sh, _, p, y in screenshot)
check("總市值", s["market_value"], want_mv, tol=0.01)
check("今日損益", s["today_pnl"], want_today, tol=0.01)
print(f"        總市值 {s['market_value']:,.0f}   今日損益 {s['today_pnl']:+,.0f}   "
      f"總損益 {s['total_pnl']:+,.0f} ({s['total_return']:+.2f}%)")

print("\n" + "=" * 60)
print(f"通過 {PASS} 項，失敗 {FAIL} 項")
print("=" * 60)
sys.exit(1 if FAIL else 0)
