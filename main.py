"""
BotShortTrade – Hyperliquid DEX
Stratégie : M15 biais → M5 confirm → M1 entrée | SL=1.5×ATR | TP=2.5×ATR | 5× levier
"""
import json, math, os, time, traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone

import requests
from dotenv import load_dotenv
from eth_account import Account
from hyperliquid.info import Info
from hyperliquid.exchange import Exchange
from hyperliquid.utils import constants
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ── Env ────────────────────────────────────────────────────────────────────────
load_dotenv()

HL_PRIVATE_KEY    = os.environ["HL_PRIVATE_KEY"]
HL_TESTNET        = os.getenv("HL_TESTNET", "1") == "1"
MAX_LIQUIDITY     = float(os.getenv("MAX_LIQUIDITY", 100))
CAPITAL_PCT       = float(os.getenv("CAPITAL_PCT", 0.10))
LEVERAGE          = int(os.getenv("LEVERAGE", 5))
MIN_MOVE_PCT      = float(os.getenv("MIN_MOVE_PCT", 0.0002))
VOL_MULTIPLIER    = float(os.getenv("VOL_MULTIPLIER", 0.60))
MAX_TRADES        = int(os.getenv("MAX_TRADES", 0))
MAX_DAILY_LOSS_PCT = float(os.getenv("MAX_DAILY_LOSS_PCT", 3.0))   # % du capital
SESSION_START_UTC  = int(os.getenv("SESSION_START_UTC", 7))         # heure UTC
SESSION_END_UTC    = int(os.getenv("SESSION_END_UTC", 22))          # heure UTC

CAPITAL_PER_TRADE = MAX_LIQUIDITY * CAPITAL_PCT   # 10 $

ATR_SL_MULT   = 1.5
ATR_TP_MULT   = 2.5
TRAIL_TRIGGER = 1.5   # activer trailing quand profit ≥ 1.5×ATR
TRAIL_DIST    = 1.0   # trailing SL à 1×ATR du prix courant

# ── Connexion Hyperliquid ──────────────────────────────────────────────────────
_account       = Account.from_key(HL_PRIVATE_KEY)
WALLET_ADDRESS = _account.address
API_URL        = constants.TESTNET_API_URL if HL_TESTNET else constants.MAINNET_API_URL
info           = Info(API_URL, skip_ws=True)
exchange       = Exchange(_account, API_URL)

print(f"  Wallet : {WALLET_ADDRESS}")
print(f"  Réseau : {'TESTNET' if HL_TESTNET else 'MAINNET'}")

# ── Symboles ───────────────────────────────────────────────────────────────────
WATCHLIST = ["BTC", "ETH", "AVAX", "LINK", "LTC", "DOGE", "XRP"]
SYM_LABEL   = {c: f"{c}/USDT" for c in WATCHLIST}
SYM_BYBIT   = {c: f"{c}USDT" for c in WATCHLIST}   # format Bybit

INTERVAL_M1  = "1m"
INTERVAL_M5  = "5m"
INTERVAL_M15 = "15m"
_BYBIT_IV    = {"1m": "1", "5m": "5", "15m": "15"}  # Bybit interval codes

BYBIT_KLINE  = "https://api.bybit.com/v5/market/kline"

# ── Filtres instruments ────────────────────────────────────────────────────────
_sym_filters: dict[str, dict] = {}

def _flt(coin: str) -> dict:
    return _sym_filters.get(coin, {"sz_dec": 3, "sz_step": 0.001})

def init_exchange_info():
    global _sym_filters
    try:
        meta = info.meta()
        for item in meta.get("universe", []):
            coin = item.get("name", "")
            if coin in set(WATCHLIST):
                sz_dec = int(item.get("szDecimals", 3))
                _sym_filters[coin] = {
                    "sz_dec":  sz_dec,
                    "sz_step": 10 ** (-sz_dec),
                    "max_lev": int(item.get("maxLeverage", 50)),
                }
        print(f"  Exchange info : {len(_sym_filters)} symboles chargés")
    except Exception as e:
        log_error(f"init_exchange_info: {e}")

# ── Excel ──────────────────────────────────────────────────────────────────────
EXCEL_FILE = "trades.xlsx"
STATE_FILE = "bot_state.json"
COL_HEADERS = [
    "ID Ordre", "Date/Heure", "Symbole", "Côté", "Qté", "Prix entrée",
    "SL", "TP", "ATR M5", "Notionnel $", "P&L TP $", "P&L SL $",
    "Prix actuel", "Statut", "P&L final $",
]

_FILLS = {
    "header": PatternFill("solid", fgColor="1F3864"),
    "buy":    PatternFill("solid", fgColor="D9EAD3"),
    "sell":   PatternFill("solid", fgColor="FCE5CD"),
    "tp":     PatternFill("solid", fgColor="B6D7A8"),
    "sl":     PatternFill("solid", fgColor="EA9999"),
    "open":   PatternFill("solid", fgColor="FFF2CC"),
}

def _safe_save(wb: Workbook):
    tmp = EXCEL_FILE + ".tmp"
    wb.save(tmp)
    os.replace(tmp, EXCEL_FILE)

def init_excel() -> Workbook:
    if os.path.exists(EXCEL_FILE):
        try:
            return openpyxl.load_workbook(EXCEL_FILE)
        except Exception:
            pass
    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"
    ws.append(COL_HEADERS)
    for c in range(1, len(COL_HEADERS) + 1):
        cell = ws.cell(1, c)
        cell.fill = _FILLS["header"]
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    for i in range(3, len(COL_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14
    _safe_save(wb)
    return wb

def _find_row(ws, order_id: str) -> int | None:
    for row in ws.iter_rows(min_row=2, values_only=False):
        if str(row[0].value) == str(order_id):
            return row[0].row
    return None

def log_order(wb: Workbook, order_id, coin, side, qty, entry, sl, tp, atr, notional, pnl_tp, pnl_sl):
    ws = wb["Trades"]
    fill = _FILLS["buy"] if side == "BUY" else _FILLS["sell"]
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
    row = [order_id, ts, SYM_LABEL[coin], side, qty, entry, sl, tp, round(atr, 6),
           notional, pnl_tp, pnl_sl, entry, "Ordre placé", ""]
    ws.append(row)
    for cell in ws[ws.max_row]:
        cell.fill = fill
    _safe_save(wb)

def update_order_live(wb: Workbook, order_id, current_px):
    ws = wb["Trades"]
    r = _find_row(ws, order_id)
    if r:
        ws.cell(r, 13).value = round(current_px, 6)
        _safe_save(wb)

def update_order_status(wb: Workbook, order_id, status: str, pnl: float | None = None):
    ws = wb["Trades"]
    r = _find_row(ws, order_id)
    if not r:
        return
    ws.cell(r, 14).value = status
    if pnl is not None:
        ws.cell(r, 15).value = round(pnl, 2)
    fill = _FILLS["tp"] if "TP" in status else (_FILLS["sl"] if "SL" in status else _FILLS["open"])
    for c in range(1, len(COL_HEADERS) + 1):
        ws.cell(r, c).fill = fill
    _safe_save(wb)

# ── Journalisation ────────────────────────────────────────────────────────────
_scan_log: list[dict] = []

def log_error(msg: str):
    ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
    print(f"  [ERR {ts}] {msg}")

def log_scan(coin: str, tf: str, signal: str | None, score: float, details: str):
    _scan_log.append({"ts": datetime.now(timezone.utc).strftime("%H:%M:%S"),
                      "coin": coin, "tf": tf, "signal": signal,
                      "score": score, "details": details})

# ── Bougies — données Bybit public (sans clé API) ────────────────────────────
def fetch_candles(coin: str, interval: str, count: int, drop_incomplete: bool = True) -> list[dict]:
    symbol = SYM_BYBIT[coin]
    iv     = _BYBIT_IV[interval]
    resp = requests.get(BYBIT_KLINE, params={
        "category": "linear",
        "symbol":   symbol,
        "interval": iv,
        "limit":    count + 2,
    }, timeout=10)
    resp.raise_for_status()
    data = resp.json()
    if data.get("retCode") != 0:
        raise RuntimeError(f"Bybit kline {symbol}: {data.get('retMsg')}")
    # Bybit retourne du plus récent au plus ancien → on inverse
    raw = data["result"]["list"]
    candles = sorted(
        [{"ts": int(r[0]), "o": float(r[1]), "h": float(r[2]),
          "l": float(r[3]), "c": float(r[4]), "v": float(r[5])}
         for r in raw],
        key=lambda x: x["ts"]
    )
    if drop_incomplete and candles:
        candles = candles[:-1]
    return candles[-count:]

# ── Indicateurs ───────────────────────────────────────────────────────────────
def _ema(closes: list[float], n: int) -> list[float]:
    k = 2 / (n + 1)
    ema = [closes[0]]
    for p in closes[1:]:
        ema.append(p * k + ema[-1] * (1 - k))
    return ema

def _atr(candles: list[dict], n: int = 14) -> float:
    trs = []
    for i in range(1, len(candles)):
        c = candles[i]
        pc = candles[i - 1]["c"]
        trs.append(max(c["h"] - c["l"], abs(c["h"] - pc), abs(c["l"] - pc)))
    if not trs:
        return 0.0
    atr = trs[0]
    k = 1 / n
    for tr in trs[1:]:
        atr = tr * k + atr * (1 - k)
    return atr

def _rsi(closes: list[float], n: int = 5) -> float:
    if len(closes) < n + 1:
        return 50.0
    gains, losses = [], []
    for i in range(-n, 0):
        d = closes[i] - closes[i - 1]
        gains.append(max(d, 0))
        losses.append(max(-d, 0))
    ag, al = sum(gains) / n, sum(losses) / n
    if al == 0:
        return 100.0
    return 100 - 100 / (1 + ag / al)

def _vwap(candles: list[dict]) -> float:
    num = sum(((c["h"] + c["l"] + c["c"]) / 3) * c["v"] for c in candles)
    den = sum(c["v"] for c in candles)
    return num / den if den else 0.0

def _avg_vol(candles: list[dict], n: int = 20) -> float:
    vols = [c["v"] for c in candles[-n:]]
    return sum(vols) / len(vols) if vols else 0.0

# ── Signal multi-timeframe ────────────────────────────────────────────────────
def get_signal(coin: str) -> dict | None:
    try:
        c15 = fetch_candles(coin, INTERVAL_M15, 30)
        c5  = fetch_candles(coin, INTERVAL_M5,  50)
        c1  = fetch_candles(coin, INTERVAL_M1,  20)
        if len(c15) < 10 or len(c5) < 20 or len(c1) < 10:
            return None

        # M15 biais (EMA8)
        ema8_15 = _ema([c["c"] for c in c15], 8)
        bias_up = c15[-1]["c"] > ema8_15[-1]
        bias_dn = c15[-1]["c"] < ema8_15[-1]

        # M5 confirmation (VWAP + EMA8 + ATR)
        closes5 = [c["c"] for c in c5]
        ema8_5  = _ema(closes5, 8)
        vwap5   = _vwap(c5)
        atr5    = _atr(c5, 14)
        last5   = closes5[-1]
        move_pct = abs(last5 - closes5[-2]) / closes5[-2] if closes5[-2] else 0

        confirm_long  = last5 > ema8_5[-1] and last5 > vwap5
        confirm_short = last5 < ema8_5[-1] and last5 < vwap5

        avg_vol5 = _avg_vol(c5)
        vol_ok   = c5[-1]["v"] >= avg_vol5 * VOL_MULTIPLIER
        move_ok  = move_pct >= MIN_MOVE_PCT

        # M1 entrée (RSI14 + volume)
        closes1  = [c["c"] for c in c1]
        rsi1     = _rsi(closes1, 14)
        avg_vol1 = _avg_vol(c1)
        vol1_ok  = c1[-1]["v"] >= avg_vol1 * VOL_MULTIPLIER

        signal = None
        score  = 0.0

        if bias_up and confirm_long and vol_ok and move_ok and rsi1 < 70 and vol1_ok:
            signal = "buy"
            score  = round((rsi1 / 100) * 0.3 + (c5[-1]["v"] / avg_vol5) * 0.4 + 0.3, 3)
        elif bias_dn and confirm_short and vol_ok and move_ok and rsi1 > 30 and vol1_ok:
            signal = "sell"
            score  = round(((100 - rsi1) / 100) * 0.3 + (c5[-1]["v"] / avg_vol5) * 0.4 + 0.3, 3)

        details = (f"M15={'↑' if bias_up else '↓'} M5={'↑' if confirm_long else '↓'} "
                   f"RSI={rsi1:.1f} vol={c5[-1]['v']:.2f}/{avg_vol5:.2f}")
        log_scan(coin, "M15/M5/M1", signal, score, details)

        if signal:
            return {"symbol": coin, "signal": signal, "score": score,
                    "atr_m5": atr5, "price": last5}
        return None

    except Exception as e:
        log_error(f"Signal {coin}: {e}")
        return None

# ── Balance (affichage uniquement) ───────────────────────────────────────────
def get_equity() -> float:
    user_state = info.user_state(WALLET_ADDRESS)
    perp_bal = float(user_state["marginSummary"]["accountValue"])
    if perp_bal > 0:
        return perp_bal
    spot = info.spot_user_state(WALLET_ADDRESS)
    return sum(float(b["total"]) for b in spot.get("balances", []) if b["coin"] == "USDC")

def get_unrealized_pnl() -> float:
    try:
        positions = info.user_state(WALLET_ADDRESS).get("assetPositions", [])
        return sum(float(p["position"]["unrealizedPnl"]) for p in positions)
    except Exception:
        return 0.0

def load_initial_balance(current_balance: float) -> float:
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE) as f:
                state = json.load(f)
            if "initial_balance" in state:
                return float(state["initial_balance"])
        except Exception:
            pass
    _save_initial_balance(current_balance)
    return current_balance

def _save_initial_balance(balance: float):
    state = {}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE) as f:
                state = json.load(f)
        except Exception:
            pass
    state["initial_balance"] = round(balance, 4)
    with open(STATE_FILE, "w") as f:
        json.dump(state, f)

# ── Scan parallèle ────────────────────────────────────────────────────────────
def scan_all(open_coins: set[str]) -> list[dict]:
    candidates = [c for c in WATCHLIST if c not in open_coins]
    results = []
    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = {ex.submit(get_signal, c): c for c in candidates}
        for fut in as_completed(futs):
            sig = fut.result()
            if sig:
                results.append(sig)
    return results

def best_signal(sigs: list[dict]) -> dict | None:
    return max(sigs, key=lambda s: s["score"]) if sigs else None

# ── Taille de position ─────────────────────────────────────────────────────────
def calc_qty(coin: str, price: float) -> float:
    f = _flt(coin)
    notional = CAPITAL_PER_TRADE * LEVERAGE
    sz = notional / price
    factor = 10 ** f["sz_dec"]
    sz = math.floor(sz * factor) / factor
    return max(f["sz_step"], sz)

def round_px(price: float) -> float:
    if price >= 1000: return round(price, 1)
    if price >= 100:  return round(price, 2)
    if price >= 10:   return round(price, 3)
    if price >= 1:    return round(price, 4)
    return round(price, 6)

# ── Levier ────────────────────────────────────────────────────────────────────
def set_leverage(coin: str):
    try:
        res = exchange.update_leverage(LEVERAGE, coin, is_cross=True)
        if res.get("status") != "ok":
            print(f"  [WARN] Levier {coin}: {res}")
    except Exception as e:
        print(f"  [WARN] Levier {coin}: {e}")

# ── Passer un ordre ───────────────────────────────────────────────────────────
def place_order(sig: dict, wb: Workbook) -> dict:
    coin   = sig["symbol"]
    atr    = sig["atr_m5"]
    is_buy = sig["signal"] == "buy"
    side   = "BUY" if is_buy else "SELL"

    set_leverage(coin)

    all_mids = info.all_mids()
    price = float(all_mids.get(coin, 0))
    if not price:
        raise RuntimeError(f"Prix indisponible: {coin}")

    qty = calc_qty(coin, price)

    if is_buy:
        sl = round_px(price - ATR_SL_MULT * atr)
        tp = round_px(price + ATR_TP_MULT * atr)
    else:
        sl = round_px(price + ATR_SL_MULT * atr)
        tp = round_px(price - ATR_TP_MULT * atr)

    print(f"  [{side}] {coin} {qty} @ ${price:.6g}  SL=${sl:.6g}  TP=${tp:.6g}  {LEVERAGE}x")

    # Ouverture de la position au marché
    result = exchange.market_open(coin, is_buy, qty, px=None, slippage=0.02)
    if result.get("status") != "ok":
        raise RuntimeError(f"Ordre refusé {coin}: {result}")

    statuses  = result["response"]["data"]["statuses"]
    fill      = statuses[0].get("filled") or statuses[0].get("resting") or {}
    order_id  = str(fill.get("oid", f"hl-{coin}-{int(time.time())}"))
    entry_px  = float(fill.get("avgPx") or price)
    open_fee  = round(float(fill.get("fee", 0)), 4)

    print(f"  [OK] oid={order_id}  entrée=${entry_px:.6g}  frais=${open_fee}")

    # Ordres SL / TP (trigger reduce-only)
    close_buy = not is_buy
    sl_lim = round_px(sl * 0.90) if is_buy else round_px(sl * 1.10)
    tp_lim = round_px(tp * 1.10) if is_buy else round_px(tp * 0.90)
    sl_oid = None

    for label, trig_px, lim_px, tpsl in [
        ("SL", sl, sl_lim, "sl"),
        ("TP", tp, tp_lim, "tp"),
    ]:
        try:
            r = exchange.order(
                coin, close_buy, qty,
                limit_px=lim_px,
                order_type={"trigger": {"triggerPx": float(trig_px), "isMarket": True, "tpsl": tpsl}},
                reduce_only=True,
            )
            if r.get("status") == "ok":
                print(f"  [{label}] ${trig_px:.6g} → OK")
                if label == "SL":
                    st = r["response"]["data"]["statuses"]
                    sl_oid = (st[0].get("resting") or st[0].get("filled") or {}).get("oid")
            else:
                print(f"  [{label}] REJETÉ : {r}")
        except Exception as e:
            print(f"  [{label}] ERREUR : {e}")

    notional = round(qty * entry_px, 2)
    pnl_tp   = round(abs(tp - entry_px) * qty * LEVERAGE, 2)
    pnl_sl   = round(-abs(sl - entry_px) * qty * LEVERAGE, 2)

    log_order(wb, order_id, coin, side, qty, entry_px, sl, tp, atr, notional, pnl_tp, pnl_sl)

    return {"id": order_id, "symbol": coin, "side": side,
            "entry": entry_px, "qty": qty, "sl": sl, "tp": tp,
            "sl_oid": sl_oid, "atr": atr, "open_fee": open_fee}

# ── Trailing stop ────────────────────────────────────────────────────────────
def trail_sl(position: dict, current_price: float):
    coin   = position["symbol"]
    is_buy = position["side"] == "BUY"
    entry  = position["entry"]
    atr    = position.get("atr", entry * 0.002)
    sl     = position["sl"]
    sl_oid = position.get("sl_oid")

    profit_atr = ((current_price - entry) / atr) if is_buy else ((entry - current_price) / atr)
    if profit_atr < TRAIL_TRIGGER:
        return

    new_sl = round_px(current_price - TRAIL_DIST * atr) if is_buy \
             else round_px(current_price + TRAIL_DIST * atr)

    # N'avancer le SL que dans le sens favorable
    if (is_buy and new_sl <= sl) or (not is_buy and new_sl >= sl):
        return

    # Annuler l'ancien SL
    if sl_oid:
        try:
            exchange.cancel(coin, int(sl_oid))
        except Exception as e:
            print(f"  [TRAIL] annulation SL : {e}")

    # Placer le nouveau SL
    close_buy = not is_buy
    sl_lim = round_px(new_sl * 0.90) if is_buy else round_px(new_sl * 1.10)
    try:
        r = exchange.order(
            coin, close_buy, position["qty"],
            limit_px=sl_lim,
            order_type={"trigger": {"triggerPx": float(new_sl), "isMarket": True, "tpsl": "sl"}},
            reduce_only=True,
        )
        if r.get("status") == "ok":
            st = r["response"]["data"]["statuses"]
            position["sl_oid"] = (st[0].get("resting") or {}).get("oid")
            position["sl"] = new_sl
            arrow = "↑" if is_buy else "↓"
            print(f"  [TRAIL {arrow}] {coin} SL={new_sl:.6g} (profit={profit_atr:.1f}×ATR)")
        else:
            print(f"  [TRAIL] rejeté : {r}")
    except Exception as e:
        print(f"  [TRAIL] erreur : {e}")

# ── Suivi de position ─────────────────────────────────────────────────────────
def check_position(wb: Workbook, position: dict) -> str | None:
    coin = position["symbol"]
    try:
        user_state = info.user_state(WALLET_ADDRESS)
        asset_pos  = user_state.get("assetPositions", [])
        pos = next(
            (p["position"] for p in asset_pos
             if p["position"]["coin"] == coin and abs(float(p["position"]["szi"])) > 0),
            None
        )

        if pos:
            mids  = info.all_mids()
            price = float(mids.get(coin, position["entry"]))
            pnl   = float(pos.get("unrealizedPnl", 0))
            pct   = (price - position["entry"]) / position["entry"] * 100
            if position["side"] == "SELL":
                pct = -pct
            print(f"  [SUIVI] {coin} ${price:.6g} ({pct:+.2f}%)  P&L=${pnl:.2f}")
            update_order_live(wb, position["id"], price)
            trail_sl(position, price)
            return None

        # Position fermée — chercher le P&L dans les fills
        pnl, result = 0.0, "Fermé"
        try:
            fills   = info.user_fills(WALLET_ADDRESS)
            closing = [f for f in fills
                       if f["coin"] == coin and "Close" in f.get("dir", "")]
            if closing:
                recent = max(closing, key=lambda f: f["time"])
                close_fee = float(recent.get("fee", 0))
                open_fee  = position.get("open_fee", close_fee)
                gross     = float(recent.get("closedPnl", 0))
                pnl       = round(gross - open_fee - close_fee, 2)
                result    = "TP touché" if pnl >= 0 else "SL touché"
                print(f"  [FRAIS] ouv=${open_fee:.4f}  ferm=${close_fee:.4f}  net=${pnl:.2f}")
        except Exception:
            pass

        print(f"  [{result}] {coin}  P&L=${pnl:.2f}")
        update_order_status(wb, position["id"], result, pnl)
        return result

    except Exception as e:
        log_error(f"check_position {coin}: {e}")
        return None

# ── Récupération positions à la reprise ──────────────────────────────────────
def recover_open_positions() -> tuple[list, float]:
    try:
        user_state = info.user_state(WALLET_ADDRESS)
        asset_pos  = user_state.get("assetPositions", [])
    except Exception as e:
        log_error(f"recover: {e}")
        return [], 0.0

    open_pos = [p["position"] for p in asset_pos
                if p["position"]["coin"] in set(WATCHLIST)
                and abs(float(p["position"]["szi"])) > 0]
    if not open_pos:
        return [], 0.0

    recovered, capital_used = [], 0.0
    for p in open_pos:
        coin  = p["coin"]
        size  = float(p["szi"])
        side  = "BUY" if size > 0 else "SELL"
        entry = float(p.get("entryPx") or 0)
        qty   = abs(size)
        atr_fb = entry * 0.002
        sl = (entry - ATR_SL_MULT * atr_fb) if side == "BUY" else (entry + ATR_SL_MULT * atr_fb)
        tp = (entry + ATR_TP_MULT * atr_fb) if side == "BUY" else (entry - ATR_TP_MULT * atr_fb)
        recovered.append({"id": f"recovered-{coin}", "symbol": coin, "side": side,
                           "entry": entry, "qty": qty, "sl": sl, "tp": tp})
        capital_used += CAPITAL_PER_TRADE
        print(f"  [REPRISE] {coin} {side} @ ${entry:.6g}")

    return recovered, capital_used

def reconcile_zombie_orders(wb: Workbook, open_positions: list) -> None:
    ws = wb["Trades"]
    open_ids = {str(p["id"]) for p in open_positions}
    for row in ws.iter_rows(min_row=2, values_only=False):
        if str(row[13].value) == "Ordre placé" and str(row[0].value) not in open_ids:
            row[13].value = "Fermé (inconnu)"
            for cell in row:
                cell.fill = _FILLS["sl"]
    _safe_save(wb)

# ── Affichage scan ────────────────────────────────────────────────────────────
# ── P&L portefeuille ──────────────────────────────────────────────────────────
def calc_portfolio_pnl(wb: Workbook) -> tuple[float, float]:
    ws = wb["Trades"]
    now = datetime.now(timezone.utc)
    total, weekly = 0.0, 0.0
    for row in ws.iter_rows(min_row=2, values_only=True):
        pnl = row[14]   # colonne O — P&L final $
        if pnl is None or pnl == "":
            continue
        try:
            pnl = float(pnl)
        except (ValueError, TypeError):
            continue
        total += pnl
        date_str = str(row[1] or "")   # colonne B — Date/Heure
        try:
            trade_dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M").replace(tzinfo=timezone.utc)
            if (now - trade_dt).days < 7:
                weekly += pnl
        except ValueError:
            pass
    return round(total, 2), round(weekly, 2)

def _print_scan_summary():
    if not _scan_log:
        return
    print(f"\n  ── Scan ({len(_scan_log)} symboles) ──")
    for e in _scan_log:
        sig = e["signal"] or "–"
        print(f"    {e['coin']:5s} {sig:4s}  score={e['score']:.3f}  {e['details']}")
    _scan_log.clear()

# ── Boucle principale ─────────────────────────────────────────────────────────
def run():
    print("\n" + "═" * 60)
    print("  BotShortTrade  –  Hyperliquid DEX")
    print("═" * 60)

    wb = init_excel()
    init_exchange_info()

    try:
        balance = get_equity()
        print(f"  Balance : ${balance:.2f} USDC")
    except Exception as e:
        log_error(f"Balance: {e}")
        balance = MAX_LIQUIDITY

    initial_balance = load_initial_balance(balance)
    print(f"  Solde initial : ${initial_balance:.2f} USDC")

    open_positions, capital_in_use = recover_open_positions()
    if open_positions:
        reconcile_zombie_orders(wb, open_positions)
        print(f"  Positions reprises : {len(open_positions)}")
    else:
        print("  Aucune position ouverte à la reprise")

    LOOP_SECONDS = 60
    print(f"\n  Démarrage — scan toutes les {LOOP_SECONDS}s\n")

    while True:
        loop_start = time.time()
        ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
        print(f"\n[{ts}] ──────────────────────────────────────────────────")

        # Suivi des positions ouvertes
        closed_this_loop = []
        for pos in list(open_positions):
            result = check_position(wb, pos)
            if result is not None:
                closed_this_loop.append(pos)

        for pos in closed_this_loop:
            open_positions.remove(pos)
            capital_in_use = max(0.0, capital_in_use - CAPITAL_PER_TRADE)

        # Balance à jour
        try:
            balance = get_equity()
        except Exception:
            pass

        # Disponible = ce qu'il reste du capital alloué (tracking interne, pas l'API)
        avail_capital = MAX_LIQUIDITY - capital_in_use
        max_slots     = MAX_TRADES if MAX_TRADES > 0 else 999
        slots_left    = max_slots - len(open_positions)

        pnl_total, pnl_week = calc_portfolio_pnl(wb)
        unrealized = get_unrealized_pnl()
        net_pnl = pnl_total + unrealized
        net_pct = (net_pnl / initial_balance * 100) if initial_balance > 0 else 0.0
        print(f"  Positions: {len(open_positions)}  Dispo: ${avail_capital:.2f}  Équité: ${balance:.2f}")
        print(f"  P&L réalisé: ${pnl_total:+.2f}  |  Non réalisé: ${unrealized:+.2f}  |  Net: ${net_pnl:+.2f} ({net_pct:+.1f}%)")
        print(f"  P&L 7j: ${pnl_week:+.2f}")

        # Filtre horaire (heures UTC)
        hour_utc = datetime.now(timezone.utc).hour
        in_session = SESSION_START_UTC <= hour_utc < SESSION_END_UTC

        # Limite de perte journalière
        _, pnl_today = calc_portfolio_pnl(wb)
        daily_loss_limit = -MAX_LIQUIDITY * MAX_DAILY_LOSS_PCT / 100
        daily_loss_hit = pnl_today <= daily_loss_limit
        if daily_loss_hit:
            print(f"  [STOP] Perte journalière atteinte (${pnl_today:.2f} ≤ ${daily_loss_limit:.2f})")
        if not in_session:
            print(f"  [PAUSE] Hors session ({hour_utc}h UTC, actif {SESSION_START_UTC}h-{SESSION_END_UTC}h)")

        # Chercher un nouveau signal
        if slots_left > 0 and avail_capital >= CAPITAL_PER_TRADE and in_session and not daily_loss_hit:
            open_coins = {p["symbol"] for p in open_positions}
            signals    = scan_all(open_coins)
            _print_scan_summary()

            sig = best_signal(signals)
            if sig:
                print(f"\n  [SIGNAL] {sig['symbol']} {sig['signal'].upper()}  score={sig['score']:.3f}")
                try:
                    pos = place_order(sig, wb)
                    open_positions.append(pos)
                    capital_in_use += CAPITAL_PER_TRADE
                except Exception as e:
                    log_error(f"place_order: {e}")
                    traceback.print_exc()
            else:
                print("  Pas de signal retenu ce cycle")
        else:
            print("  Slots/capital épuisés — scan ignoré")
            _scan_log.clear()

        elapsed = time.time() - loop_start
        wait    = max(0.0, LOOP_SECONDS - elapsed)
        if wait > 0:
            time.sleep(wait)


if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        print("\n  [STOP] Arrêt demandé")
    except Exception as e:
        log_error(f"Crash fatal: {e}")
        traceback.print_exc()
