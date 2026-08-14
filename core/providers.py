"""報價與匯率來源。

2026-08-14 實測結論（決定了這裡的每一個選擇）：
  * 台股走證交所 MIS，需先取 cookie 並帶 Referer，否則被擋。
    13:30 後的成交價 z 就是收盤價，所以不需要另外做盤後校正
    （證交所 STOCK_DAY_ALL 當天下午還是前一交易日的資料，不能用）。
  * 美股走 Yahoo chart API。**必須用 range=1d**，因為昨收要讀 meta.chartPreviousClose，
    而該欄位在 range=5d 時給的是「區間開始前」的收盤價，會把漲跌算成完全錯的數字。
    meta.regularMarketPreviousClose 這個欄位不存在，別找它。
  * yfinance 1.6.0 的 fast_info 全部回 None，只有 history() 能用，所以只拿它當備援。
  * 台灣銀行牌告匯率被反爬蟲擋死（回驗證頁），改用 Yahoo USDTWD=X + open.er-api.com。
  * Stooq 全部 404，已棄用。
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field

import requests

from . import config as C


@dataclass
class Quote:
    code: str
    name: str = ""
    price: float | None = None
    prev_close: float | None = None
    currency: str = "TWD"
    source: str = ""
    stale: bool = False          # True = 這筆是靠昨收或買賣價遞補出來的
    error: str = ""
    data_date: str = ""          # 來源回報的資料日期 YYYYMMDD，用來判斷是否休市

    @property
    def ok(self) -> bool:
        return self.price is not None


@dataclass
class FxRate:
    rate: float | None = None
    source: str = ""
    note: str = ""
    error: str = ""


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": C.USER_AGENT,
                      "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8"})
    return s


def _num(v):
    """MIS 的數字欄位可能是 '-'、''、'101.5000'。"""
    try:
        if v is None:
            return None
        t = str(v).strip()
        if t in ("", "-", "--"):
            return None
        return float(t)
    except (TypeError, ValueError):
        return None


# ══════════════════════════════════════════════════════════════════
#  台股：證交所 MIS
# ══════════════════════════════════════════════════════════════════
class TwseMisProvider:
    """證交所基本市況報導的資料源，盤中約每 5 秒更新。"""

    INDEX = "https://mis.twse.com.tw/stock/index.jsp"
    API = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
    BATCH = 40                       # 一次查太多檔會被拒絕

    def __init__(self):
        self._s: requests.Session | None = None
        self._otc: set[str] = set()  # 已知走上櫃的代號，下次直接用 otc_

    def _ensure_session(self):
        if self._s is None:
            s = _session()
            s.headers["Referer"] = self.INDEX
            s.get(self.INDEX, timeout=C.HTTP_TIMEOUT)   # 取 cookie，少了會被擋
            self._s = s
        return self._s

    def _fetch(self, codes: list[str], prefix: str) -> dict[str, dict]:
        if not codes:
            return {}
        s = self._ensure_session()
        out: dict[str, dict] = {}
        for i in range(0, len(codes), self.BATCH):
            chunk = codes[i:i + self.BATCH]
            r = s.get(self.API, timeout=C.HTTP_TIMEOUT, params={
                "ex_ch": "|".join(f"{prefix}_{c}.tw" for c in chunk),
                "json": "1", "delay": "0", "_": str(int(time.time() * 1000)),
            })
            r.raise_for_status()
            for it in (r.json() or {}).get("msgArray", []) or []:
                if it.get("c"):
                    out[it["c"]] = it
        return out

    def get_quotes(self, codes: list[str]) -> dict[str, Quote]:
        if not codes:
            return {}
        results: dict[str, Quote] = {}
        try:
            # 已知上櫃的直接走 otc_，其餘先試 tse_
            otc_first = [c for c in codes if c in self._otc]
            tse_first = [c for c in codes if c not in self._otc]

            raw = self._fetch(tse_first, "tse")
            raw.update(self._fetch(otc_first, "otc"))

            # tse_ 查不到的，改用 otc_ 再試一次（上櫃股票）
            missing = [c for c in codes if c not in raw]
            if missing:
                found = self._fetch(missing, "otc")
                self._otc.update(found.keys())
                raw.update(found)
        except Exception as e:
            return {c: Quote(code=c, error=f"證交所 MIS 連線失敗：{e}") for c in codes}

        for code in codes:
            it = raw.get(code)
            if not it:
                results[code] = Quote(code=code, error="證交所查無此代號")
                continue

            price = _num(it.get("z"))
            stale = False
            if price is None:
                # 無成交時用最佳買價／賣價中值遞補，再不行就用昨收
                bid = _num((it.get("b") or "").split("_")[0])
                ask = _num((it.get("a") or "").split("_")[0])
                if bid and ask:
                    price, stale = (bid + ask) / 2, True
                elif bid or ask:
                    price, stale = (bid or ask), True
                else:
                    price, stale = _num(it.get("y")), True

            results[code] = Quote(
                code=code,
                name=(it.get("n") or "").strip(),
                price=price,
                prev_close=_num(it.get("y")),
                currency="TWD",
                source="證交所 MIS",
                stale=stale,
                error="" if price is not None else "無成交價可用",
                data_date=str(it.get("d") or "").strip(),
            )
        return results


# ══════════════════════════════════════════════════════════════════
#  Yahoo chart API：美股主力、台股備援、匯率
# ══════════════════════════════════════════════════════════════════
class YahooChartProvider:
    URL = "https://query1.finance.yahoo.com/v8/finance/chart/{sym}"

    def __init__(self):
        self._s = _session()

    def get_meta(self, symbol: str) -> dict | None:
        # range=1d 是關鍵：只有這樣 chartPreviousClose 才等於真正的昨收
        r = self._s.get(self.URL.format(sym=symbol), timeout=C.HTTP_TIMEOUT,
                        params={"range": "1d", "interval": "1d"})
        r.raise_for_status()
        res = (r.json().get("chart") or {}).get("result") or []
        return res[0].get("meta") if res else None

    def get_quote(self, code: str, symbol: str | None = None) -> Quote:
        sym = symbol or code
        try:
            m = self.get_meta(sym)
            if not m:
                return Quote(code=code, error=f"Yahoo 查無 {sym}")
            return Quote(
                code=code,
                name=(m.get("shortName") or m.get("longName") or "").strip(),
                price=m.get("regularMarketPrice"),
                prev_close=m.get("chartPreviousClose"),
                currency=m.get("currency") or "USD",
                source="Yahoo Finance",
            )
        except Exception as e:
            return Quote(code=code, error=f"Yahoo 連線失敗：{e}")

    def get_quotes(self, codes: list[str], suffix: str = "") -> dict[str, Quote]:
        return {c: self.get_quote(c, f"{c}{suffix}") for c in codes}


class YFinanceProvider:
    """備援。1.6.0 的 fast_info 是壞的，只用 history()。"""

    def get_quotes(self, codes: list[str], suffix: str = "") -> dict[str, Quote]:
        out: dict[str, Quote] = {}
        try:
            import yfinance as yf
        except Exception as e:
            return {c: Quote(code=c, error=f"yfinance 無法載入：{e}") for c in codes}

        for c in codes:
            try:
                df = yf.Ticker(f"{c}{suffix}").history(period="5d")
                closes = [float(x) for x in df["Close"].tolist() if x == x]   # 濾掉 NaN
                if not closes:
                    out[c] = Quote(code=c, error="yfinance 無資料")
                    continue
                out[c] = Quote(
                    code=c,
                    price=closes[-1],
                    prev_close=closes[-2] if len(closes) >= 2 else None,
                    currency="USD" if not suffix else "TWD",
                    source="yfinance (備援)",
                )
            except Exception as e:
                out[c] = Quote(code=c, error=f"yfinance 失敗：{e}")
        return out


# ══════════════════════════════════════════════════════════════════
#  匯率
# ══════════════════════════════════════════════════════════════════
class FxProvider:
    ER_API = "https://open.er-api.com/v6/latest/USD"

    def __init__(self, yahoo: YahooChartProvider | None = None):
        self._yahoo = yahoo or YahooChartProvider()
        self._s = _session()

    def get_usdtwd(self, manual: float | None = None) -> FxRate:
        if manual:
            return FxRate(rate=float(manual), source="手動指定",
                          note="在 持股.xlsx 的「設定」分頁指定")

        try:
            m = self._yahoo.get_meta("USDTWD=X")
            if m and m.get("regularMarketPrice"):
                return FxRate(rate=float(m["regularMarketPrice"]),
                              source="Yahoo USDTWD=X", note="市場中間價")
        except Exception:
            pass

        try:
            j = self._s.get(self.ER_API, timeout=C.HTTP_TIMEOUT).json()
            rate = (j.get("rates") or {}).get("TWD")
            if rate:
                return FxRate(rate=float(rate), source="open.er-api.com",
                              note="市場中間價，一天更新一次")
        except Exception as e:
            return FxRate(error=f"匯率抓取失敗：{e}")

        return FxRate(error="匯率抓取失敗：所有來源都沒有回應")


# ══════════════════════════════════════════════════════════════════
#  對外統一入口
# ══════════════════════════════════════════════════════════════════
class QuoteService:
    def __init__(self):
        self.mis = TwseMisProvider()
        self.yahoo = YahooChartProvider()
        self.yf = YFinanceProvider()
        self.fx = FxProvider(self.yahoo)

    def fetch(self, tw_codes: list[str], us_codes: list[str],
              manual_fx: float | None = None) -> tuple[dict[str, Quote], FxRate]:
        quotes: dict[str, Quote] = {}

        # ── 台股：MIS 為主，失敗的個別改用 Yahoo ──
        if tw_codes:
            quotes.update(self.mis.get_quotes(tw_codes))
            failed = [c for c in tw_codes if not quotes[c].ok]
            if failed:
                for c, q in self.yahoo.get_quotes(failed, suffix=".TW").items():
                    if q.ok:
                        q.source = "Yahoo (備援)"
                        quotes[c] = q
                # .TW 不行的再試 .TWO（上櫃）
                still = [c for c in failed if not quotes[c].ok]
                for c, q in self.yahoo.get_quotes(still, suffix=".TWO").items():
                    if q.ok:
                        q.source = "Yahoo 上櫃 (備援)"
                        quotes[c] = q

        # ── 美股：Yahoo 為主，失敗的改用 yfinance ──
        if us_codes:
            quotes.update(self.yahoo.get_quotes(us_codes))
            failed = [c for c in us_codes if not quotes[c].ok]
            if failed:
                for c, q in self.yf.get_quotes(failed).items():
                    if q.ok:
                        quotes[c] = q

        return quotes, self.fx.get_usdtwd(manual_fx)
