"""重複代號的行為測試。

規則：允許重複（分批買進想分開追蹤），但存檔前要先提醒；
      重複的每一筆各自保留股數與均價，不合併，總計加總。
"""
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook

from core import config as C

TMP = Path(tempfile.mkdtemp(prefix="finance_dup_"))
C.HOLDINGS_XLSX = TMP / "持股.xlsx"
C.DATA_DIR = TMP / "data"
C.SETTINGS_JSON = C.DATA_DIR / "設定.json"
C.HISTORY_XLSX = C.DATA_DIR / "歷史紀錄.xlsx"
C.EXPORT_XLSX = C.DATA_DIR / "庫存總覽.xlsx"
C.BACKUP_DIR = C.DATA_DIR / "備份"
C.CACHE_JSON = C.DATA_DIR / "報價快取.json"

from core import excel_io                                   # noqa: E402
from core.excel_io import Holding                           # noqa: E402
from core.portfolio import build_rows, summarize            # noqa: E402
from core.providers import FxRate, Quote                    # noqa: E402
import app as app_module                                    # noqa: E402

FAIL = 0


def ok(name, cond, extra=""):
    global FAIL
    if cond:
        print(f"  PASS  {name}" + (f"  {extra}" if extra else ""))
    else:
        FAIL += 1
        print(f"  FAIL  {name}  {extra}")


def make(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = C.SHEET_HOLDINGS
    ws.append(C.HOLDING_COLUMNS)
    for r in rows:
        ws.append(r)
    wb.save(C.HOLDINGS_XLSX)


# ══════════════════════════════════════════════════════════════════
print("=" * 74)
print("1. 讀檔：重複不再是錯誤，而是警告")
print("=" * 74)
make([
    ["2882", "國泰金", "TW", 100, 40.0, "第一批"],
    ["00919", "", "TW", 5000, 22.0, ""],
    ["2882", "國泰金", "TW", 200, 50.0, "第二批"],
    ["2882", "國泰金", "TW", 300, 60.0, "第三批"],
])
book = excel_io.read_holdings()
ok("不再丟例外，四列都讀進來", len(book.holdings) == 4, f"{len(book.holdings)} 列")
ok("重複清單正確", book.duplicates == {"2882": [2, 4, 5]}, str(book.duplicates))
ok("有警告訊息", any("重複" in w for w in book.warnings), str(book.warnings))
ok("警告有提到會分開計算", any("分開" in w for w in book.warnings))

dup_rows = [h for h in book.holdings if h.code == "2882"]
ok("三筆各自保留自己的股數", [h.shares for h in dup_rows] == [100, 200, 300],
   str([h.shares for h in dup_rows]))
ok("三筆各自保留自己的均價", [h.avg_cost for h in dup_rows] == [40.0, 50.0, 60.0])
ok("序號標成 1/3、2/3、3/3",
   [(h.dup_index, h.dup_total) for h in dup_rows] == [(1, 3), (2, 3), (3, 3)],
   str([(h.dup_index, h.dup_total) for h in dup_rows]))
ok("備註沒有被搞混", [h.note for h in dup_rows] == ["第一批", "第二批", "第三批"])

single = next(h for h in book.holdings if h.code == "00919")
ok("沒重複的那檔序號是 0", (single.dup_index, single.dup_total) == (0, 0))

make([["2882", "", "TW", 100, 40.0, ""], ["00919", "", "TW", 5000, 22.0, ""]])
book = excel_io.read_holdings()
ok("沒有重複時不產生警告", not book.duplicates and not book.warnings, str(book.warnings))


# ══════════════════════════════════════════════════════════════════
print("\n" + "=" * 74)
print("2. 計算：分開算，總計加總")
print("=" * 74)
holdings = [
    Holding(code="2882", name="國泰金", market=C.MARKET_TW, shares=100, avg_cost=40.0),
    Holding(code="2882", name="國泰金", market=C.MARKET_TW, shares=200, avg_cost=50.0),
]
excel_io.mark_duplicates(holdings)
quotes = {"2882": Quote(code="2882", price=100.0, prev_close=90.0)}
rows = build_rows(holdings, quotes, FxRate(rate=32.0))

ok("兩列都在", len(rows) == 2)
ok("第一筆市值 100x100", rows[0]["市值"] == 10000.0)
ok("第二筆市值 200x100", rows[1]["市值"] == 20000.0)
ok("第一筆成本 100x40", rows[0]["成本"] == 4000.0)
ok("第二筆成本 200x50", rows[1]["成本"] == 10000.0)
ok("兩筆報酬率不同（沒有被合併）",
   abs(rows[0]["報酬率"] - 150.0) < 1e-9 and abs(rows[1]["報酬率"] - 100.0) < 1e-9,
   f'{rows[0]["報酬率"]:.1f}% / {rows[1]["報酬率"]:.1f}%')
ok("兩筆共用同一份報價", rows[0]["現價"] == rows[1]["現價"] == 100.0)
ok("序號傳到畫面資料",
   (rows[0]["重複序號"], rows[0]["重複總數"]) == (1, 2)
   and (rows[1]["重複序號"], rows[1]["重複總數"]) == (2, 2))

s = summarize(rows, FxRate(rate=32.0))
ok("總市值加總 = 30000", s["market_value"] == 30000.0)
ok("總成本加總 = 14000", s["cost"] == 14000.0)
ok("總損益 = 16000", s["total_pnl"] == 16000.0)
ok("今日損益加總 = 300x10", s["today_pnl"] == 3000.0)

print("\n  對照：同樣部位填成一列（加權平均 46.6667）應得到相同總計")
merged = [Holding(code="2882", market=C.MARKET_TW, shares=300,
                  avg_cost=14000 / 300, name="")]
ms = summarize(build_rows(merged, quotes, FxRate(rate=32.0)), FxRate(rate=32.0))
ok("總市值一致", abs(ms["market_value"] - s["market_value"]) < 1e-6)
ok("總成本一致", abs(ms["cost"] - s["cost"]) < 1e-6)
ok("總報酬率一致", abs(ms["total_return"] - s["total_return"]) < 1e-9,
   f'{ms["total_return"]:.4f}% vs {s["total_return"]:.4f}%')

print("\n  無報價的重複檔，代號只列一次不重複")
h2 = [Holding(code="9999", market=C.MARKET_TW, shares=1, avg_cost=1, name=""),
      Holding(code="9999", market=C.MARKET_TW, shares=2, avg_cost=1, name="")]
excel_io.mark_duplicates(h2)
s2 = summarize(build_rows(h2, {"9999": Quote(code="9999", error="查無")}, FxRate(rate=32.0)),
               FxRate(rate=32.0))
ok("incomplete 不重複列出", s2["incomplete"] == ["9999"], str(s2["incomplete"]))


# ══════════════════════════════════════════════════════════════════
print("\n" + "=" * 74)
print("3. 存檔：先提醒，確認後才寫")
print("=" * 74)
api = app_module.Api()
payload = [
    {C.COL_CODE: "2882", C.COL_SHARES: 100, C.COL_AVGCOST: 40, C.COL_MARKET: "TW"},
    {C.COL_CODE: "2882", C.COL_SHARES: 200, C.COL_AVGCOST: 50, C.COL_MARKET: "TW"},
    {C.COL_CODE: "0050", C.COL_SHARES: 10, C.COL_AVGCOST: 100, C.COL_MARKET: "TW"},
]

before = excel_io.read_holdings().holdings
res = api.save_holdings(payload)
ok("第一次呼叫不存檔，要求確認", res.get("needs_confirm") == "duplicates", str(res))
ok("回傳重複的列號", res.get("duplicates") == {"2882": [1, 2]}, str(res.get("duplicates")))
ok("檔案沒有被改動", [h.code for h in excel_io.read_holdings().holdings]
   == [h.code for h in before])

res = api.save_holdings(payload, confirmed=True)
ok("確認後真的存進去", res.get("ok"), str(res))
ok("存檔結果附帶重複資訊", res.get("duplicates") == {"2882": [1, 2]})

saved = excel_io.read_holdings()
ok("讀回來是 3 列", len(saved.holdings) == 3, f"{len(saved.holdings)} 列")
kept = [(h.code, h.shares, h.avg_cost) for h in saved.holdings]
ok("兩筆 2882 都在且數字沒被合併",
   kept == [("2882", 100.0, 40.0), ("2882", 200.0, 50.0), ("0050", 10.0, 100.0)], str(kept))

print("\n  沒有重複時不該要求確認:")
res = api.save_holdings([{C.COL_CODE: "2330", C.COL_SHARES: 10, C.COL_AVGCOST: 900}])
ok("直接存成功", res.get("ok") and not res.get("needs_confirm"), str(res))
ok("重複清單是空的", res.get("duplicates") == {})

print("\n  真正的錯誤仍然要擋（優先於重複提醒）:")
res = api.save_holdings([
    {C.COL_CODE: "2882", C.COL_SHARES: 100, C.COL_AVGCOST: 40},
    {C.COL_CODE: "2882", C.COL_SHARES: "abc", C.COL_AVGCOST: 40},
])
ok("有錯就回 errors，不是 needs_confirm",
   res.get("errors") and not res.get("needs_confirm"), str(res)[:90])
ok("錯誤指向股數欄", res["errors"][0]["field"] == C.COL_SHARES)

res = api.save_holdings([
    {C.COL_CODE: "", C.COL_SHARES: 100, C.COL_AVGCOST: 40},
    {C.COL_CODE: "2882", C.COL_SHARES: 100, C.COL_AVGCOST: 40},
    {C.COL_CODE: "2882", C.COL_SHARES: 100, C.COL_AVGCOST: 40},
], confirmed=True)
ok("空白代號仍然被擋", res.get("errors") and res["errors"][0]["field"] == C.COL_CODE)


# ══════════════════════════════════════════════════════════════════
print("\n" + "=" * 74)
print("4. 報價查詢去重")
print("=" * 74)
make([
    ["2882", "", "TW", 100, 40, ""],
    ["2882", "", "TW", 200, 50, ""],
    ["2882", "", "TW", 300, 60, ""],
    ["AAPL", "", "US", 10, 100, ""],
    ["AAPL", "", "US", 20, 200, ""],
])

calls = {"tw": [], "us": []}
real_fetch = api.service.fetch


def spy(tw, us, manual_fx=None):
    calls["tw"], calls["us"] = list(tw), list(us)
    quotes = {c: Quote(code=c, price=100.0, prev_close=90.0,
                       currency="USD" if c == "AAPL" else "TWD", source="test")
              for c in list(tw) + list(us)}
    return quotes, FxRate(rate=32.0, source="test")


api.service.fetch = spy
payload = api.refresh()
api.service.fetch = real_fetch

ok("台股代號只送一次", calls["tw"] == ["2882"], str(calls["tw"]))
ok("美股代號只送一次", calls["us"] == ["AAPL"], str(calls["us"]))
ok("畫面仍然是 5 列", len(payload["rows"]) == 5, f'{len(payload["rows"])} 列')
ok("五列都有價格", all(r["現價"] is not None for r in payload["rows"]))
ok("警告有提到重複", any("重複" in w for w in payload["warnings"]), str(payload["warnings"]))

tw_total = 100 * 100 + 200 * 100 + 300 * 100
us_total = (10 + 20) * 100 * 32.0
ok("總市值 = 各列加總", abs(payload["summary"]["market_value"] - (tw_total + us_total)) < 0.01,
   f'{payload["summary"]["market_value"]:,.0f}')


shutil.rmtree(TMP, ignore_errors=True)
print("\n" + "=" * 74)
print("全部通過" if not FAIL else f"有 {FAIL} 項失敗")
print("=" * 74)
sys.exit(1 if FAIL else 0)
