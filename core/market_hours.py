"""判斷現在是不是盤中，決定自動更新的節奏。

收盤後價格不會變，一直打證交所 API 沒意義還可能被限速。
"""
from __future__ import annotations

import datetime as dt
from zoneinfo import ZoneInfo

from . import config as C

_TPE = ZoneInfo(C.TZ_TAIPEI)
_NYC = ZoneInfo(C.TZ_NEWYORK)


def _in_window(now: dt.datetime, start: tuple[int, int], end: tuple[int, int]) -> bool:
    if now.weekday() >= 5:                      # 週六日一定沒開
        return False
    t = now.time()
    return dt.time(*start) <= t <= dt.time(*end)


def tw_open(now: dt.datetime | None = None) -> bool:
    return _in_window((now or dt.datetime.now(_TPE)).astimezone(_TPE),
                      C.TW_OPEN, C.TW_CLOSE)


def us_open(now: dt.datetime | None = None) -> bool:
    """用紐約當地時間判斷，日光節約由 zoneinfo 自動處理。"""
    return _in_window((now or dt.datetime.now(_NYC)).astimezone(_NYC),
                      C.US_OPEN, C.US_CLOSE)


def next_interval(has_tw: bool, has_us: bool, *,
                  tw_data_is_today: bool | None = None,
                  now: dt.datetime | None = None) -> tuple[int, str]:
    """回傳 (秒數, 給使用者看的說明)。

    tw_data_is_today: 證交所回來的資料日期是不是今天。
        遇到颱風假、國定假日時，時間看起來在盤中但實際沒開市，
        這個旗標會把節奏拉回非盤中，不會空打 API。
    """
    now = now or dt.datetime.now(_TPE)
    reasons: list[str] = []

    tw = has_tw and tw_open(now)
    if tw and tw_data_is_today is False:
        tw = False
        reasons.append("台股今日休市")
    elif tw:
        reasons.append("台股盤中")

    us = has_us and us_open(now)
    if us:
        reasons.append("美股盤中")

    if tw or us:
        return C.REFRESH_OPEN_SEC, "、".join(reasons)

    if not reasons:
        reasons.append("非交易時段")
    return C.REFRESH_CLOSED_SEC, "、".join(reasons)
