"""持股管理 —— 程式進入點。

用 pywebview 開一個獨立的 Windows 視窗，介面是 HTML，運算是 Python。
JS 透過 window.pywebview.api.* 直接呼叫下面 Api 類別的方法，中間沒有網頁伺服器。
"""
from __future__ import annotations

import datetime
import os
import subprocess
import sys
import traceback
from pathlib import Path

import webview

from core import config as C
from core import excel_io, market_hours, portfolio
from core import settings as settings_store
from core.excel_io import HoldingsFileError
from core.providers import FxRate, QuoteService


def _tw_data_is_today(quotes: dict, tw_codes: list[str]) -> bool | None:
    """證交所回報的資料日期是不是今天。

    遇到颱風假或國定假日時，時鐘看起來在盤中但市場其實沒開，
    靠這個判斷把自動更新拉回較慢的節奏，不會空打 API。
    回傳 None 表示無從判斷（例如報價來自備援來源）。
    """
    today = datetime.date.today().strftime("%Y%m%d")
    dates = [quotes[c].data_date for c in tw_codes
             if c in quotes and quotes[c].data_date]
    if not dates:
        return None
    return any(d == today for d in dates)


class Api:
    def __init__(self):
        self.service = QuoteService()

    # ── 啟動時：先把快取畫出來，不要讓使用者對著空畫面等網路 ──────────
    def get_initial(self) -> dict:
        created = excel_io.ensure_holdings_template()
        migrated = settings_store.migrate_from_excel()
        payload = {
            "created_template": created,
            "holdings_path": str(C.HOLDINGS_XLSX),
            "data_dir": str(C.DATA_DIR),
            "settings": settings_store.load(),
        }
        if created:
            payload["notice"] = (
                f"第一次執行，已幫你建立範本檔：\n{C.HOLDINGS_XLSX}\n\n"
                f"按「編輯持股」把範例改成你自己的持股。"
            )
        elif migrated:
            payload["notice"] = (
                f"設定已從 {C.HOLDINGS_XLSX.name} 搬到 {C.SETTINGS_JSON.name}，"
                f"該檔的「設定」分頁已移除。\n"
                f"以後改設定請按右上角的「設定」按鈕，不用開 Excel。"
            )
        cached = excel_io.load_cache()
        if cached:
            cached["from_cache"] = True
            payload["cached"] = cached
        return payload

    # ── 主要動作：讀 Excel → 抓報價 → 算損益 ──────────────────────
    def refresh(self) -> dict:
        conf = settings_store.load()
        try:
            book = excel_io.read_holdings()
        except HoldingsFileError as e:
            return {"error": str(e), "error_kind": "holdings", "settings": conf}
        except Exception as e:
            return {"error": f"讀取持股檔案時發生非預期錯誤：\n{e}",
                    "error_kind": "holdings", "settings": conf}

        if not book.holdings:
            return {"error": "持股清單是空的。\n按「編輯持股」新增你的第一筆持股。",
                    "error_kind": "empty", "settings": conf}

        tw = [h.code for h in book.holdings if h.market == C.MARKET_TW]
        us = [h.code for h in book.holdings if h.market == C.MARKET_US]

        try:
            quotes, fx = self.service.fetch(tw, us, conf.get(C.SET_MANUAL_FX))
        except Exception as e:
            traceback.print_exc()
            return {"error": f"抓取報價時發生錯誤：\n{e}",
                    "error_kind": "network", "settings": conf}

        rows = portfolio.build_rows(book.holdings, quotes, fx)
        summary = portfolio.summarize(rows, fx)

        interval, why = market_hours.next_interval(
            has_tw=bool(tw), has_us=bool(us),
            tw_data_is_today=_tw_data_is_today(quotes, tw))

        payload = {
            "rows": rows,
            "summary": summary,
            "warnings": book.warnings,
            "settings": conf,
            "holdings_mtime": book.mtime,
            # interval 是「現在這個時段該用的節奏」，關掉自動更新時仍要送，
            # 設定視窗才能告訴使用者打開後會是多久更新一次
            "auto_interval_sec": interval,
            "auto_refresh_sec": interval if conf.get(C.SET_AUTO_REFRESH) else 0,
            "market_state": why,
            "holdings_path": str(C.HOLDINGS_XLSX),
        }

        # 每日淨值只在算得完整時才記錄，免得把殘缺的數字寫進歷史
        try:
            if not summary["incomplete"] and summary["market_value"] > 0:
                excel_io.append_nav_record(portfolio.nav_record(summary))
                payload["nav_saved"] = True
        except Exception as e:
            payload["warnings"] = list(payload["warnings"]) + [f"歷史淨值寫入失敗：{e}"]

        try:
            excel_io.save_cache(payload)
        except Exception:
            pass
        return payload

    # ── 編輯持股 ──────────────────────────────────────────────────
    def get_holdings_for_edit(self) -> dict:
        """進入編輯模式時呼叫，回傳目前 Excel 裡的原始內容。"""
        try:
            book = excel_io.read_holdings()
        except HoldingsFileError as e:
            return {"error": str(e)}
        except Exception as e:
            return {"error": f"讀取持股檔案失敗：\n{e}"}

        return {
            "rows": [{C.COL_CODE: h.code, C.COL_NAME: h.name, C.COL_MARKET: h.market,
                      C.COL_SHARES: h.shares, C.COL_AVGCOST: h.avg_cost,
                      C.COL_NOTE: h.note} for h in book.holdings],
            "mtime": book.mtime,
            "excel_open": excel_io.excel_lock_holder(C.HOLDINGS_XLSX) is not None,
        }

    def lookup(self, code: str, market: str = "") -> dict:
        """編輯時輸入代號，即時查名稱與現價，讓使用者確認沒打錯。"""
        code = (code or "").strip()
        if not code:
            return {"ok": False}

        mkt = market.strip().upper() if market else excel_io.infer_market(code)
        try:
            if mkt == C.MARKET_TW:
                quotes, _ = self.service.fetch([code], [], manual_fx=1.0)
            else:
                quotes, _ = self.service.fetch([], [code], manual_fx=1.0)
        except Exception as e:
            return {"ok": False, "error": str(e)}

        q = quotes.get(code)
        if not q or not q.ok:
            return {"ok": False, "market": mkt,
                    "error": (q.error if q else "") or "查無此代號"}
        return {"ok": True, "market": mkt, "name": q.name, "price": q.price,
                "currency": q.currency, "source": q.source}

    def save_holdings(self, rows, expect_mtime=None) -> dict:
        """把編輯結果寫回 持股.xlsx。驗證不過就不寫，並指出是哪一列。"""
        cleaned, errors = [], []
        seen: dict[str, int] = {}

        for i, r in enumerate(rows or [], start=1):
            code = str(r.get(C.COL_CODE, "") or "").strip()
            if not code:
                errors.append({"index": i - 1, "field": C.COL_CODE, "msg": "代號不能空白"})
                continue
            if code in seen:
                errors.append({"index": i - 1, "field": C.COL_CODE,
                               "msg": f"代號重複，第 {seen[code]} 列已經有了"})
                continue
            seen[code] = i

            try:
                shares = float(str(r.get(C.COL_SHARES, "")).replace(",", "").strip())
                if shares <= 0:
                    raise ValueError
            except (TypeError, ValueError):
                errors.append({"index": i - 1, "field": C.COL_SHARES, "msg": "股數要是大於 0 的數字"})
                continue

            try:
                avg = float(str(r.get(C.COL_AVGCOST, "")).replace(",", "").replace("$", "").strip())
                if avg < 0:
                    raise ValueError
            except (TypeError, ValueError):
                errors.append({"index": i - 1, "field": C.COL_AVGCOST, "msg": "均價要是 0 或正數"})
                continue

            market = str(r.get(C.COL_MARKET, "") or "").strip().upper()
            if market not in (C.MARKET_TW, C.MARKET_US):
                market = excel_io.infer_market(code)

            cleaned.append({
                C.COL_CODE: code,
                C.COL_NAME: str(r.get(C.COL_NAME, "") or "").strip(),
                C.COL_MARKET: market,
                C.COL_SHARES: shares,
                C.COL_AVGCOST: avg,
                C.COL_NOTE: str(r.get(C.COL_NOTE, "") or "").strip(),
            })

        if errors:
            return {"ok": False, "errors": errors}

        try:
            info = excel_io.write_holdings(cleaned, expect_mtime=expect_mtime)
        except HoldingsFileError as e:
            return {"ok": False, "error": str(e)}
        except Exception as e:
            traceback.print_exc()
            return {"ok": False, "error": f"存檔失敗：\n{e}"}

        return {"ok": True, **info}

    # ── 設定 ──────────────────────────────────────────────────────
    def get_settings(self) -> dict:
        return settings_store.load()

    def save_settings(self, values: dict) -> dict:
        try:
            return {"ok": True, "settings": settings_store.save(values or {})}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ── 使用者操作 ────────────────────────────────────────────────
    def export_overview(self, rows, summary) -> dict:
        try:
            path = excel_io.export_overview(rows, summary)
            return {"ok": True, "path": str(path)}
        except PermissionError:
            return {"ok": False, "error": f"寫不進 {C.EXPORT_XLSX.name}，"
                                          f"請先關掉正在開著它的 Excel 再試一次。"}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_holdings(self) -> dict:
        return self._open(C.HOLDINGS_XLSX)

    def open_data_dir(self) -> dict:
        C.DATA_DIR.mkdir(parents=True, exist_ok=True)
        return self._open(C.DATA_DIR)

    @staticmethod
    def _open(path: Path) -> dict:
        try:
            if sys.platform == "win32":
                os.startfile(str(path))                       # noqa: S606
            else:
                subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(path)])
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}


def _work_area() -> tuple[int, int]:
    """取得扣掉工作列後的可用桌面大小。

    刻意不用 webview.screens —— 在 webview.start() 之前存取它會提早初始化
    GUI 後端，導致視窗被建成 159x27 這種壞尺寸（開發時踩到過）。
    """
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            rect = wintypes.RECT()
            SPI_GETWORKAREA = 0x0030
            if ctypes.windll.user32.SystemParametersInfoW(
                    SPI_GETWORKAREA, 0, ctypes.byref(rect), 0):
                return rect.right - rect.left, rect.bottom - rect.top
            return (ctypes.windll.user32.GetSystemMetrics(0),
                    ctypes.windll.user32.GetSystemMetrics(1))
        except Exception:
            pass
    return 1536, 816


def _window_size() -> tuple[int, int]:
    """視窗大小跟著螢幕走。寫死尺寸在小螢幕或高 DPI 縮放下會超出可視範圍。"""
    sw, sh = _work_area()
    return (max(720, min(1440, int(sw * 0.92))),
            max(520, min(900, int(sh * 0.94))))


def _report_fatal(exc: BaseException) -> None:
    """用 pythonw 啟動時沒有主控台，錯誤會無聲無息地消失。
    所以把完整 traceback 寫進 data\\error.log，再跳一個視窗告訴使用者去哪看。
    """
    detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    log = C.DATA_DIR / "error.log"
    try:
        C.DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(log, "a", encoding="utf-8") as f:
            f.write(f"\n{'=' * 70}\n{datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n{detail}")
    except Exception:
        pass

    print(detail, file=sys.stderr)
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0,
                f"持股管理啟動失敗。\n\n{type(exc).__name__}: {exc}\n\n"
                f"完整錯誤已寫入：\n{log}",
                "持股管理",
                0x10,                                  # MB_ICONERROR
            )
        except Exception:
            pass


def main() -> None:
    C.DATA_DIR.mkdir(parents=True, exist_ok=True)
    excel_io.ensure_holdings_template()

    w, h = _window_size()
    api = Api()
    webview.create_window(
        "持股管理",
        str(C.UI_DIR / "index.html"),
        js_api=api,
        width=w,
        height=h,
        min_size=(720, 520),
        maximized=True,
        background_color="#12151C",
    )
    webview.start()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        _report_fatal(e)
        sys.exit(1)
