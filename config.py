# ═══════════════════════════════════════════════════════════════════════════════
#  config.py  —  shared settings for all three NSE Volume Profile scripts
# ═══════════════════════════════════════════════════════════════════════════════

# ── Symbols ───────────────────────────────────────────────────────────────────
SYMBOLS = [
    "RELIANCE", "TCS", "INFY", "HDFCBANK", "ICICIBANK",
    "HINDUNILVR", "AXISBANK", "KOTAKBANK", "SBIN", "BHARTIARTL",
]

# ── Profile parameters ────────────────────────────────────────────────────────
LOOKBACK_DAYS   = 30    # daily sessions for volume_profile.py
LOOKBACK_WEEKS  = 12    # weeks for weekly_volume_profile.py
PRICE_BINS      = 20    # bins per session / week
VALUE_AREA_PCT  = 70    # standard 70% value area

# ── S/R clustering ────────────────────────────────────────────────────────────
DAILY_TOLERANCE_PCT   = 0.5   # price bins within 0.5% → same daily cluster
WEEKLY_TOLERANCE_PCT  = 0.6   # slightly wider for weekly
DAILY_MIN_TOUCHES     = 2     # sessions to qualify as daily S/R
DAILY_STRONG_TOUCHES  = 4     # sessions for STRONG daily label
WEEKLY_MIN_TOUCHES    = 2     # weeks to qualify as weekly S/R
WEEKLY_STRONG_TOUCHES = 3     # weeks for STRONG weekly label
SR_TOP_N              = 3     # max zones per side (3 R + 3 S) in summary sheets

# ── Files ─────────────────────────────────────────────────────────────────────
DAILY_EXCEL   = "volume_profile_daily.xlsx"    # output of volume_profile.py
SR_EXCEL      = "volume_profile_sr.xlsx"       # output of add_sr_levels.py
WEEKLY_EXCEL  = "volume_profile_weekly.xlsx"   # output of weekly_volume_profile.py

# ── Google Sheets ─────────────────────────────────────────────────────────────
# Set GOOGLE_SHEET_ID to enable Sheets push; leave "" to skip.
GOOGLE_SHEET_ID    = ""                    # e.g. "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74..."
GOOGLE_CREDS_JSON  = "service_account.json"
SHEET_SR_ZONES     = "SR Zones"           # tab for daily S/R (written by add_sr_levels.py)
SHEET_WEEKLY_SR    = "Weekly SR Zones"    # tab for weekly S/R (written by weekly_volume_profile.py)

# ── Telegram ──────────────────────────────────────────────────────────────────
# Set both to enable Telegram; leave "" to skip.
TELEGRAM_BOT_TOKEN = ""    # e.g. "123456789:ABCdef..."
TELEGRAM_CHAT_ID   = ""    # e.g. "-1001234567890"

# ── Proximity alert threshold (Oracle Cloud alert script) ─────────────────────
ALERT_BAND_PCT = 0.5    # fire alert when LTP is within 0.5% of VAH or VAL
