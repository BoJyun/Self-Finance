"""路徑、常數與 Excel 結構定義。"""
from pathlib import Path

# ── 路徑 ────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent

HOLDINGS_XLSX = ROOT / "持股.xlsx"        # 使用者的檔案，程式只讀不寫
DATA_DIR = ROOT / "data"                  # 程式的檔案，使用者平常不用開
HISTORY_XLSX = DATA_DIR / "歷史紀錄.xlsx"
EXPORT_XLSX = DATA_DIR / "庫存總覽.xlsx"
BACKUP_DIR = DATA_DIR / "備份"
CACHE_JSON = DATA_DIR / "報價快取.json"
UI_DIR = ROOT / "ui"

KEEP_BACKUPS = 30                         # 備份保留份數

# ── 持股.xlsx 結構 ──────────────────────────────────────────────────
SHEET_HOLDINGS = "持股"

COL_CODE = "代號"
COL_NAME = "名稱"
COL_MARKET = "市場"
COL_SHARES = "股數"
COL_AVGCOST = "均價"
COL_NOTE = "備註"

HOLDING_COLUMNS = [COL_CODE, COL_NAME, COL_MARKET, COL_SHARES, COL_AVGCOST, COL_NOTE]
REQUIRED_COLUMNS = [COL_CODE, COL_SHARES, COL_AVGCOST]   # 這三欄缺一不可

MARKET_TW = "TW"
MARKET_US = "US"

# ── 設定 ────────────────────────────────────────────────────────────
# 設定存在 data\設定.json（程式偏好），不放在 持股.xlsx（你的資產資料）。
# 理由：選了「程式可寫回 Excel」之後，Excel 開著時寫入會失敗；
#      不該因為 Excel 開著就連換個顏色都做不到。
SETTINGS_JSON = DATA_DIR / "設定.json"

SET_MANUAL_FX = "手動匯率"
SET_COLOR_SCHEME = "漲跌顏色"
SET_AUTO_REFRESH = "自動更新"

COLOR_TW_STYLE = "紅漲綠跌"
COLOR_US_STYLE = "綠漲紅跌"

DEFAULT_SETTINGS = {
    SET_MANUAL_FX: None,          # None = 自動抓匯率
    SET_COLOR_SCHEME: COLOR_TW_STYLE,
    SET_AUTO_REFRESH: True,       # 盤中 60 秒、非盤中 10 分鐘
}

# 舊版把設定放在 持股.xlsx 的「設定」分頁，首次啟動時搬過來後移除該分頁
SHEET_SETTINGS_LEGACY = "設定"
LEGACY_AUTO_REFRESH = "自動更新秒數"

# ── 自動更新節奏 ────────────────────────────────────────────────────
REFRESH_OPEN_SEC = 60          # 盤中
REFRESH_CLOSED_SEC = 600       # 非盤中

TZ_TAIPEI = "Asia/Taipei"
TZ_NEWYORK = "America/New_York"
TW_OPEN, TW_CLOSE = (9, 0), (13, 30)        # 台股交易時間
US_OPEN, US_CLOSE = (9, 30), (16, 0)        # 美股一般交易時段（當地時間）

# ── 歷史紀錄結構 ────────────────────────────────────────────────────
SHEET_NAV = "每日淨值"
NAV_COLUMNS = ["日期", "總市值", "總成本", "總損益", "報酬率", "台股市值",
               "美股市值(台幣)", "USD/TWD"]

# ── 網路 ────────────────────────────────────────────────────────────
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
HTTP_TIMEOUT = 20
