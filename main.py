"""
BotShortTrade – Hyperliquid DEX
Stratégie tendance journalière sur BTC, ETH, SOL, XRP et DOGE : vote 14/28/56 j, SL 3×ATR D1
(cf. section « Stratégie tendance journalière » et tools/trend_backtest.py).
L'ancien scalping M1 (aucun avantage sur le hasard après frais) a été retiré le 2026-11-11 ; il reste
dans l'historique git (étiquette scalping-final).
"""
import json, math, os, sys, time, traceback
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

# ── Couleurs console (ANSI) ────────────────────────────────────────────────────
os.system("")  # active les codes ANSI sur Windows

class C:
    RST  = "\033[0m"
    BOLD = "\033[1m"
    DIM  = "\033[2m"
    RED  = "\033[91m"
    GRN  = "\033[92m"
    YLW  = "\033[93m"
    BLU  = "\033[94m"
    CYN  = "\033[96m"
    WHT  = "\033[97m"
    BRED = "\033[1;91m"
    BGRN = "\033[1;92m"
    BYLW = "\033[1;93m"
    BCYN = "\033[1;96m"

def _cp(val: float) -> str:
    """Couleur selon signe d'une valeur P&L."""
    return C.BGRN if val > 0 else (C.BRED if val < 0 else C.WHT)

# ── Env ────────────────────────────────────────────────────────────────────────
load_dotenv()

HL_PRIVATE_KEY    = os.environ["HL_PRIVATE_KEY"]
HL_TESTNET        = os.getenv("HL_TESTNET", "1") == "1"
_MAX_LIQ_CFG      = float(os.getenv("MAX_LIQUIDITY", 0))   # plafond de l'équité servant aux tailles (0 = tout le solde)
LEVERAGE          = int(os.getenv("LEVERAGE", 5))          # levier croisé réglé avant chaque ouverture
# Mode fantôme : aucun ordre réel, positions tendance suivies au prix moyen (validation de la mécanique)
GHOST_ONLY        = os.getenv("GHOST_ONLY", "0") == "1"
# Capital fictif du mode fantôme (test avant le réel) : positions fantômes calculées comme si le compte
# avait ce montant (0 = solde réel). Ignoré en réel.
PAPER_EQUITY      = float(os.getenv("PAPER_EQUITY", 0))
TREND_EXPOSURE    = float(os.getenv("TREND_EXPOSURE", 0.5))   # notionnel total / équité, réparti sur TREND_COINS
DISCORD_WEBHOOK_URL = os.getenv("DISCORD_WEBHOOK_URL", "")          # webhook Discord (alertes anomalies)
EVENTS_FILE = "discord_events.jsonl"   # copie locale de chaque message Discord (relue par tools/status.py)
HEARTBEAT_HOURS    = float(os.getenv("HEARTBEAT_HOURS", 6))          # résumé Discord périodique (0 = désactivé)

LOOP_SECONDS     = 60        # un cycle par minute
LEDGER_REFRESH_S = 10 * 60   # relevé des dépôts / retraits (journal Hyperliquid)
SL_CHECK_S       = 5 * 60    # vérification que chaque position a bien son stop-loss sur Hyperliquid
STARTUP_RETRY_S  = 30        # équité illisible au démarrage : nouvel essai après ce délai

# Stratégie tendance (cf. section « Stratégie tendance journalière » et tools/trend_backtest.py)
TREND_LOOKBACKS    = (14, 28, 56)  # LONG si clôture D1 > celle d'il y a N jours pour 2 horizons sur 3
TREND_SL_ATR       = 3.0        # SL = 3×ATR14 journalier
TREND_FEE          = 0.000432   # frais taker par côté (P&L des positions fantômes)
TREND_MIN_NOTIONAL = 10.5       # $ minimum de CHAQUE ordre, à son pire prix (Hyperliquid refuse < 10 $)
TREND_MIN_TARGET   = TREND_MIN_NOTIONAL * 1.01   # visé 1 % plus haut : arrondi du prix limite par le SDK,
                                                 # prix qui bouge entre sa lecture et l'ordre
TREND_SLIP_OPEN    = 0.02       # prix limite d'un ordre au marché : 2 % (ouverture, renfort)…
TREND_SLIP_CLOSE   = 0.05       # … 5 % (fermeture)
TREND_SLIP_SL      = 0.10       # … 10 % au-delà du déclenchement (stop-loss)
TREND_COINS        = ["BTC", "ETH", "SOL", "XRP", "DOGE"]   # BTC a remplacé LTC le 2026-11-03

# ── Connexion Hyperliquid ──────────────────────────────────────────────────────
_account       = Account.from_key(HL_PRIVATE_KEY)
WALLET_ADDRESS = _account.address
API_URL        = constants.TESTNET_API_URL if HL_TESTNET else constants.MAINNET_API_URL
# Délai max d'une requête Hyperliquid. Le SDK n'en met aucun : une connexion restée bloquée (coupure
# Wi-Fi) figeait le bot sans erreur ni alerte. Ordre sans réponse : cf. « opening » (tendance).
HL_TIMEOUT_S   = 20
info           = Info(API_URL, skip_ws=True, timeout=HL_TIMEOUT_S)
exchange       = Exchange(_account, API_URL, timeout=HL_TIMEOUT_S)

print(f"  {C.DIM}Wallet : {WALLET_ADDRESS}{C.RST}")
_net_color = C.BYLW if HL_TESTNET else C.BRED
print(f"  Réseau : {_net_color}{'TESTNET' if HL_TESTNET else 'MAINNET'}{C.RST}")
if GHOST_ONLY:
    print(f"  Mode   : {C.BYLW}FANTÔME SEUL — aucun ordre réel (GHOST_ONLY=1){C.RST}")

# ── Bougies journalières (Bybit, sans clé API) ─────────────────────────────────
INTERVAL_D1  = "1D"
_BYBIT_IV    = {"1D": "D"}   # code d'intervalle Bybit
BYBIT_KLINE  = "https://api.bybit.com/v5/market/kline"

# ── Filtres instruments ────────────────────────────────────────────────────────
_sym_filters: dict[str, dict] = {}
_sym_filters_ok = True   # False après un échec de chargement : tailles inconnues, aucune ouverture réelle

def _flt(coin: str) -> dict:
    return _sym_filters.get(coin, {"sz_dec": 3, "sz_step": 0.001})

def init_exchange_info() -> bool:
    """Charge le pas de taille (szDecimals) de tous les coins de Hyperliquid : une position sur un coin
    retiré de TREND_COINS doit encore pouvoir être renforcée puis fermée. False si Hyperliquid ne répond pas."""
    global _sym_filters, _sym_filters_ok
    try:
        meta = info.meta()
        for item in meta.get("universe", []):
            sz_dec = int(item.get("szDecimals", 3))
            _sym_filters[item.get("name", "")] = {"sz_dec": sz_dec, "sz_step": 10 ** (-sz_dec)}
        print(f"  Exchange info : {len(_sym_filters)} symboles chargés")
        _sym_filters_ok = True
    except Exception as e:
        log_error(f"init_exchange_info: {e}")
        _sym_filters_ok = False
    return _sym_filters_ok

# ── Excel (trades.xlsx) ────────────────────────────────────────────────────────
# Onglet « Tendance » : une ligne par position. Les onglets de l'ancien scalping (« Trades »,
# « Fantômes ») restent dans le fichier comme historique, sans être modifiés.
EXCEL_FILE = "trades.xlsx"
STATE_FILE = "bot_state.json"   # référence du « % depuis le départ » et dernier relevé des dépôts

_FILLS = {
    "header": PatternFill("solid", fgColor="1F3864"),
    "tp":     PatternFill("solid", fgColor="B6D7A8"),   # position fermée en gain
    "sl":     PatternFill("solid", fgColor="EA9999"),   # position fermée en perte
}

TREND_HEADERS = [
    "ID", "Ouvert (UTC)", "Symbole", "Côté", "Mode", "Qté", "Entrée", "SL",
    "Notionnel $", "Statut", "Sortie", "P&L $", "P&L %", "Fermé (UTC)",
]

def _ensure_sheet(wb: Workbook, name: str, headers: list[str]) -> bool:
    """Crée la feuille `name` avec ses en-têtes si absente. Retourne True si créée."""
    if name in wb.sheetnames:
        return False
    ws = wb.create_sheet(name)
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.fill = _FILLS["header"]
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions["A"].width = 24
    return True

def _safe_save(wb: Workbook):
    tmp = EXCEL_FILE + ".tmp"
    wb.save(tmp)
    os.replace(tmp, EXCEL_FILE)

def init_excel() -> Workbook:
    """trades.xlsx, avec son onglet « Tendance » (créé s'il manque)."""
    if os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            if _ensure_sheet(wb, "Tendance", TREND_HEADERS):
                _safe_save(wb)
            return wb
        except Exception as e:
            # Fichier illisible : mis de côté avant d'en créer un neuf (avant : écrasé, historique perdu)
            bad = f"trades-illisible-{datetime.now(timezone.utc):%Y%m%d-%H%M%S}.xlsx"
            try:
                os.replace(EXCEL_FILE, bad)
            except OSError as err:
                log_error(f"{EXCEL_FILE} : mise de côté impossible ({err})")
            log_error(f"{EXCEL_FILE} illisible ({e}) : copie gardée sous {bad}, nouveau fichier créé")
            send_discord_alert(f"⚠️ {EXCEL_FILE} illisible",
                               f"Copie gardée sous `{bad}` ; un nouveau fichier est créé. Les positions "
                               f"et les SL ne sont pas touchés.\nErreur : `{str(e)[:300]}`", color=0xf39c12)
    wb = Workbook()
    wb.remove(wb.active)   # feuille vide créée par défaut
    _ensure_sheet(wb, "Tendance", TREND_HEADERS)
    _safe_save(wb)
    return wb

def log_error(msg: str):
    ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
    print(f"  {C.BRED}[ERR {ts}]{C.RST} {msg}")

# ── Bougies — données Bybit public (sans clé API) ────────────────────────────
def fetch_candles(coin: str, interval: str, count: int, drop_incomplete: bool = True) -> list[dict]:
    symbol   = f"{coin}USDT"
    iv       = _BYBIT_IV[interval]
    last_err: Exception | None = None
    for attempt in range(3):   # 2 retries (limite de débit ou panne passagère de Bybit)
        try:
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
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(0.5 + attempt)   # 0.5 s puis 1.5 s avant de réessayer
    raise last_err  # type: ignore[misc]

# ── Indicateur ────────────────────────────────────────────────────────────────
def _atr(candles: list[dict], n: int = 14) -> float:
    """ATR de Wilder initialisé sur la SMA des n premières TR (convergence immédiate).
    Un ATR mal initialisé fausse directement le niveau du SL."""
    trs = []
    for i in range(1, len(candles)):
        c  = candles[i]
        pc = candles[i - 1]["c"]
        trs.append(max(c["h"] - c["l"], abs(c["h"] - pc), abs(c["l"] - pc)))
    if not trs:
        return 0.0
    if len(trs) < n:
        return sum(trs) / len(trs)
    k   = 1 / n
    atr = sum(trs[:n]) / n   # seed = SMA(n) — convergence immédiate
    for tr in trs[n:]:
        atr = tr * k + atr * (1 - k)
    return atr

# ── Équité du compte ──────────────────────────────────────────────────────────
_unified: bool | None = None   # compte en mode Unified Hyperliquid ? (réglage du compte, lu une fois)

def _unified_account() -> bool:
    """Mode Unified Hyperliquid (le USDC spot sert de marge perp) ? Lu une fois par démarrage ; lève une
    exception si l'API ne répond pas (rien n'est mémorisé : nouvel essai au prochain appel)."""
    global _unified
    if _unified is None:
        _unified = info.query_user_abstraction_state(WALLET_ADDRESS) == "unifiedAccount"
    return _unified

def get_equity(cached_state: dict | None = None) -> float:
    """Lève une exception si l'API ne répond pas (502/504…) : l'appelant garde alors le
    solde précédent. Ne jamais remplacer une partie manquante par 0 (faux "-100 %")."""
    user_state = cached_state if cached_state is not None else info.user_state(WALLET_ADDRESS)
    cross        = user_state.get("crossMarginSummary") or user_state.get("marginSummary")
    if not cross:
        raise ValueError("user_state sans marginSummary")
    acv          = float(cross.get("accountValue") or 0)   # perp : marge engagée + P&L latent (+ libre)

    spot_usdc = 0.0
    spot_state = info.spot_user_state(WALLET_ADDRESS)   # en échec → exception propagée
    for b in spot_state.get("balances", []):
        if b["coin"] == "USDC":
            spot_usdc = float(b.get("total") or 0)
            break

    # Mode Unified : le USDC spot sert de marge perp, et son total contient déjà la marge ET le P&L latent
    # des positions ; c'est la valeur du compte donnée par Hyperliquid (endpoint portfolio, relevés du 06/11).
    # Avant : accountValue perp moins la marge ajouté en plus → P&L latent compté environ 2,2 fois.
    # Compte classique : perp et spot sont séparés → somme des deux.
    total = spot_usdc if _unified_account() else acv + spot_usdc
    return total if total > 0 else 0.0

def _sizing_equity(balance: float) -> float:
    """Équité sur laquelle la stratégie tendance taille ses positions : en mode fantôme, le capital fictif
    PAPER_EQUITY s'il est fixé (test) ; sinon tout le solde, relu à chaque cycle (dépôts et gains compris,
    comme le backtest), ou au plus MAX_LIQUIDITY s'il est fixé dans .env. Avant : capital figé au
    démarrage, recalé seulement après un écart de plus de 20 %."""
    if GHOST_ONLY and PAPER_EQUITY > 0:
        return PAPER_EQUITY
    return min(balance, _MAX_LIQ_CFG) if _MAX_LIQ_CFG > 0 else balance

def _since_start(balance: float, initial_balance: float) -> str:
    """« +1.2 % depuis le départ » : équité comparée au solde de référence (bot_state.json)."""
    pct = (balance / initial_balance - 1) * 100 if initial_balance > 0 else 0.0
    return f"{round(pct, 1) + 0.0:+.1f} % depuis le départ"   # + 0.0 : jamais « -0.0 % »

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
    state = _read_state()
    state["initial_balance"] = round(balance, 4)
    _state_atomic_write(state)

def _read_state() -> dict:
    """Contenu de bot_state.json ({} s'il manque ou est illisible)."""
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except Exception:
        return {}

# Lignes du journal Hyperliquid sans mouvement de fonds externe (le reste non reconnu est signalé)
_LEDGER_INTERNAL = {"accountClassTransfer", "liquidation", "rewardsClaim", "spotGenesis"}

def _ledger_flow(d: dict) -> float:
    """USDC entré (+) ou sorti (−) du compte selon une ligne du journal Hyperliquid : dépôt, retrait,
    virement depuis ou vers une autre adresse, coffre (vault). 0 pour le reste : passage spot ↔ perp
    du compte lui-même, autre jeton que l'USDC (hors de l'équité suivie)."""
    me, kind = WALLET_ADDRESS.lower(), d.get("type")
    if kind == "deposit":
        return float(d.get("usdc") or 0)
    if kind == "withdraw":   # frais (1 $) laissés en perte : inclus ou non dans « usdc », jamais un gain
        return -float(d.get("usdc") or 0)
    if kind == "vaultDeposit":
        return -float(d.get("usdc") or 0)
    if kind == "vaultWithdraw":
        return float(d.get("netWithdrawnUsd") or d.get("usdc") or 0)
    if kind in ("send", "spotTransfer", "internalTransfer", "subAccountTransfer"):
        if d.get("token", "USDC") != "USDC":
            return 0.0
        amount = float(d.get("usdcValue") or d.get("amount") or d.get("usdc") or 0)
        src, dst = str(d.get("user", "")).lower(), str(d.get("destination", "")).lower()
        if dst == me and src != me:
            return amount
        if src == me and dst != me:
            return -amount - float(d.get("fee") or 0)
        return 0.0
    if kind not in _LEDGER_INTERNAL:
        log_error(f"journal Hyperliquid : opération « {kind} » non comptée comme dépôt ni retrait")
    return 0.0

def sync_deposits() -> list[tuple[int, float]]:
    """Dépôts (+) et retraits (−) de USDC faits depuis le dernier relevé du journal Hyperliquid, ajoutés à
    la référence du « % depuis le départ » (bot_state.json) : un dépôt n'est pas un gain, ni un retrait
    une perte. Une alerte par mouvement. Retourne les mouvements (heure en ms, montant) ; aucun si rien
    de nouveau ou si Hyperliquid ne répond pas (nouvel essai au relevé suivant)."""
    state = _read_state()
    if "ledger_ts" not in state:   # premier relevé : les mouvements passés sont déjà dans la référence
        state["ledger_ts"] = int(time.time() * 1000)
        _state_atomic_write(state)
        return []
    try:
        rows = info.user_non_funding_ledger_updates(WALLET_ADDRESS, int(state["ledger_ts"]) + 1)
    except Exception as e:
        log_error(f"journal des dépôts Hyperliquid (nouvel essai plus tard) : {e}")
        return []
    if not rows:
        return []
    ref, moves = float(state.get("initial_balance") or 0), []
    for r in sorted(rows, key=lambda r: r["time"]):
        flow = _ledger_flow(r.get("delta") or {})
        if not flow:
            continue
        moves.append((int(r["time"]), flow))
        why = "un dépôt n'est pas un gain" if flow > 0 else "un retrait n'est pas une perte"
        send_discord_alert(
            f"💰 Dépôt de ${flow:.2f} reçu" if flow > 0 else f"💸 Retrait de ${-flow:.2f}",
            ("Utilisé dès les prochaines positions (ouvertures, inversions) ; celles déjà ouvertes gardent "
             "leur taille." if flow > 0 else "Les prochaines positions seront plus petites.")
            + f"\nRéférence du « % depuis le départ » : ${ref:.2f} → **${ref + flow:.2f}** ({why}).",
            color=0x3498db)
        ref += flow
    state["ledger_ts"] = max(int(r["time"]) for r in rows)
    if moves:
        state["initial_balance"] = round(ref, 4)
    _state_atomic_write(state)
    return moves

def _state_atomic_write(state: dict):
    """Écriture atomique de bot_state.json via tmp+replace (même pattern que _safe_save)."""
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w") as f:
        json.dump(state, f, indent=2)
    os.replace(tmp, STATE_FILE)

def round_px(price: float, coin: str | None = None) -> float:
    """Prix accepté par Hyperliquid (perps) : 5 chiffres significatifs au plus et au plus
    6 − szDecimals décimales (un prix entier est toujours valide). Avant : décimales fixes
    par tranche de prix → 0.123456 (DOGE ≥ 0.1, XRP < 1) avait 6 chiffres → SL refusé."""
    if price <= 0:
        return price
    decimals = 4 - math.floor(math.log10(price))          # 5 chiffres significatifs
    sz_dec   = _sym_filters.get(coin, {}).get("sz_dec", 0) if coin else 0
    return round(price, max(0, min(decimals, 6 - sz_dec)))

# ── Levier ────────────────────────────────────────────────────────────────────
def set_leverage(coin: str):
    try:
        res = exchange.update_leverage(LEVERAGE, coin, is_cross=True)
        if res.get("status") != "ok":
            print(f"  {C.BYLW}[WARN]{C.RST} Levier {coin}: {res}")
    except Exception as e:
        print(f"  {C.BYLW}[WARN]{C.RST} Levier {coin}: {e}")

# ── Alertes Discord ───────────────────────────────────────────────────────────
def send_discord_alert(title: str, description: str, color: int = 0xe74c3c):
    """Envoie une alerte embed sur Discord via webhook. Silencieux si pas configuré.
    Chaque message est aussi ajouté à EVENTS_FILE pour pouvoir le relire localement."""
    try:
        with open(EVENTS_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps({"t": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                                "title": title, "description": description},
                               ensure_ascii=False) + "\n")
    except Exception as e:
        log_error(f"journal Discord: {e}")
    if not DISCORD_WEBHOOK_URL:
        return
    payload = {
        "embeds": [{
            "title": title,
            "description": description,
            "color": color,
            "footer": {"text": f"BotShortTrade · {WALLET_ADDRESS[:6]}…{WALLET_ADDRESS[-4:]}"},
            "timestamp": datetime.now(timezone.utc).isoformat()
        }]
    }
    for _ in range(3):
        try:
            r = requests.post(DISCORD_WEBHOOK_URL, json=payload, timeout=5)
            if r.status_code != 429:
                r.raise_for_status()
                if r.headers.get("X-RateLimit-Remaining") == "0":   # le message suivant serait refusé
                    time.sleep(min(float(r.headers.get("X-RateLimit-Reset-After") or 1), 5))
                return
            # Limite de débit du webhook (≈ 5 messages / 2 s) : rafale de minuit ou du passage en réel.
            # Avant : message perdu sans bruit.
            time.sleep(min(float(r.headers.get("Retry-After") or 1), 5))
        except Exception:
            return  # Ne jamais crasher le bot à cause de Discord

# ── Surveillance des stop-loss ───────────────────────────────────────────────
def ensure_sl_orders(open_positions: list) -> int:
    """Vérifie que chaque position suivie a bien un ordre stop-loss actif sur Hyperliquid
    (reprise en « état partiel », SL annulé à la main, échec silencieux…). Repose le SL
    manquant au niveau position['sl'] et resynchronise sl_oid. Retourne le nb de SL reposés."""
    if not open_positions:
        return 0
    try:
        orders = info.frontend_open_orders(WALLET_ADDRESS)
    except Exception as e:
        log_error(f"ensure_sl_orders: {e}")
        return 0
    replaced = 0
    for pos in open_positions:
        coin   = pos["symbol"]
        is_buy = pos["side"] == "BUY"
        # Un stop du même sens que la position (« B » = achat) ne la protège pas : SL orphelin d'une
        # position inverse précédente, dont l'annulation est en attente
        stops  = [o for o in orders if o.get("coin") == coin and o.get("isTrigger")
                  and o.get("reduceOnly") and str(o.get("orderType", "")).startswith("Stop")
                  and o.get("side") != ("B" if is_buy else "A")]
        if stops:
            if not any(o.get("oid") == pos.get("sl_oid") for o in stops):
                pos["sl_oid"] = stops[0]["oid"]   # resynchroniser (ex. reprise état partiel)
            continue
        sl     = round_px(pos["sl"], coin)
        sl_lim = round_px(sl * 0.90, coin) if is_buy else round_px(sl * 1.10, coin)
        try:
            r = exchange.order(coin, not is_buy, pos["qty"], limit_px=sl_lim,
                               order_type={"trigger": {"triggerPx": float(sl), "isMarket": True, "tpsl": "sl"}},
                               reduce_only=True)
            st = (r.get("response") or {}).get("data", {}).get("statuses", [{}])
            if r.get("status") != "ok" or "error" in st[0]:
                raise RuntimeError(r)
            pos["sl_oid"] = (st[0].get("resting") or st[0].get("filled") or {}).get("oid")
            replaced += 1
            print(f"  {C.BYLW}[SL MANQUANT]{C.RST} {coin} SL reposé à ${sl:.6g}")
            send_discord_alert(f"🛡️ SL manquant reposé — {coin}",
                               f"Aucun stop-loss actif trouvé sur Hyperliquid ; SL reposé à **${sl:.6g}**.",
                               color=0xf39c12)
        except Exception as e:
            log_error(f"SL manquant {coin} non reposé : {e}")
            send_discord_alert(f"🚨 {coin} SANS STOP-LOSS",
                               f"SL absent sur Hyperliquid et impossible à reposer. **Intervention manuelle requise.**\n"
                               f"Erreur : `{str(e)[:300]}`", color=0xff0000)
    return replaced

# ── Fills de fermeture (P&L réel d'une position fermée) ───────────────────────
def _closing_fills(position: dict, fills: list) -> list:
    """Fills de l'ordre de fermeture le plus récent de la position ([] si aucun). Un ordre de
    fermeture (SL ou au marché) est souvent exécuté en plusieurs fills partiels : tous ceux de cet
    ordre (même oid), pas seulement le dernier."""
    coin    = position["symbol"]
    since   = (position.get("opened_at") or 0) * 1000 - 5000   # marge horloge 5 s
    # Uniquement les fermetures postérieures à l'ouverture : sinon la fermeture d'un trade
    # précédent sur le même coin pouvait être attribuée à celui-ci
    closing = [f for f in fills if f["coin"] == coin and "Close" in f.get("dir", "")
               and f.get("time", 0) >= since]
    if not closing:
        return []
    last_oid = max(closing, key=lambda f: f["time"]).get("oid")
    return [f for f in closing if f.get("oid") == last_oid]

def _closing_pnl(position: dict, fills: list) -> tuple[float, float, float] | None:
    """(gross, open_fee, close_fee) de la fermeture la plus récente de la position, ou None."""
    order = _closing_fills(position, fills)
    if not order:
        return None
    gross     = sum(float(f.get("closedPnl", 0)) for f in order)
    close_fee = sum(float(f.get("fee", 0)) for f in order)
    # Frais d'ouverture réels via les fills de l'ordre d'entrée (absents de la réponse d'ordre)
    opening  = [f for f in fills if str(f.get("oid")) == str(position.get("id"))]
    open_fee = (sum(float(f.get("fee", 0)) for f in opening) if opening
                else position.get("open_fee") or close_fee)   # estimé = frais de sortie (même taux taker)
    return gross, open_fee, close_fee

# ── Stratégie tendance journalière ───────────────────────────────────────────
# Momentum de série temporelle : chaque jour, pour chaque coin, LONG si la clôture D1 est
# au-dessus de celle d'il y a 14, 28 et 56 jours pour au moins 2 de ces horizons (vote), SHORT
# sinon. Position tenue jusqu'à l'inversion du signal ; SL à TREND_SL_ATR×ATR14 D1, puis à plat
# jusqu'à la clôture journalière suivante. Notionnel = TREND_EXPOSURE × équité / nb de coins, et aucun
# ordre sous TREND_MIN_NOTIONAL à son pire prix (fermeture trop petite : renfort d'abord, _trend_top_up).
# Backtest (tools/trend_backtest.py, TREND_COINS, 50 % de l'équité) : 2017→2026 +50 %/an, Sharpe
# 1.11, pire baisse −45 % ; 2e moitié (2022→2026) +27 %/an, Sharpe 0.99, pire baisse −34 %.
# Le vote bat la règle 28 j seule sur chaque moitié (écart non significatif, t=0.9) ; BTC à la
# place de LTC aussi (+6 %/an, t=2.2). Hypothèses déclarées avant chaque test (docstring du backtest).
# GHOST_ONLY : positions fantômes au prix moyen, aucun ordre (validation de la mécanique).
TREND_FILE = "trend_state.json"
# {"positions": {coin: position}, "stopped": {coin: jour UTC},
#  "opening": {coin: ouverture sans réponse de Hyperliquid}, "cancels": [[coin, oid] de SL à annuler]}
_trend: dict | None = None
_trend_sig: dict[str, tuple] = {}      # coin → (jour UTC, côté, clôture D1, ATR14 D1)
_trend_alerted: set[tuple] = set()     # (coin, jour, motif) : une seule alerte d'échec par jour
_trend_sl_ts = 0.0                     # dernière vérification des SL tendance sur Hyperliquid

def _utc_day(ts: float) -> str:
    return datetime.fromtimestamp(ts, timezone.utc).strftime("%Y-%m-%d")

def _today() -> str:
    return _utc_day(time.time())

def _load_trend() -> dict:
    global _trend
    if _trend is None:
        try:
            with open(TREND_FILE, encoding="utf-8") as f:
                _trend = json.load(f)
        except Exception:
            _trend = {}
        _trend.setdefault("positions", {})
        _trend.setdefault("stopped", {})
        _trend.setdefault("opening", {})
        _trend.setdefault("cancels", [])
    return _trend

def _save_trend():
    tmp = TREND_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(_load_trend(), f, indent=2)
    os.replace(tmp, TREND_FILE)

def _hl_d1(coin: str, count: int) -> list[dict]:
    """Les `count` dernières bougies D1 closes de Hyperliquid, au format de fetch_candles. Secours
    quand Bybit ne répond pas ou tarde : clôtures à 0.05 % près en moyenne et mêmes votes (vérifié le
    2026-11-04)."""
    end = int(time.time()) // 86400 * 86_400_000          # minuit UTC : bougie du jour (non close) exclue
    raw = info.candles_snapshot(coin, "1d", end - (count + 1) * 86_400_000, end - 1)
    cs  = [{"ts": int(c["t"]), "o": float(c["o"]), "h": float(c["h"]), "l": float(c["l"]),
            "c": float(c["c"]), "v": float(c["v"])} for c in raw]
    return sorted(cs, key=lambda c: c["ts"])[-count:]

def trend_signal(coin: str) -> tuple[str, float, float] | None:
    """(côté, dernière clôture D1, ATR14 D1) calculé sur les bougies journalières closes, en cache
    pour la journée UTC : Bybit, ou Hyperliquid si Bybit ne répond pas ou n'a pas la bougie d'hier.
    None si aucune source ne l'a (juste après minuit, ou panne) : nouvel essai au cycle suivant."""
    day = _today()
    cached = _trend_sig.get(coin)
    if cached and cached[0] == day:
        return cached[1:]
    n = max(TREND_LOOKBACKS)
    yesterday = (int(time.time()) // 86400 - 1) * 86_400_000
    errors = []
    for source, fetch in (("Bybit", lambda: fetch_candles(coin, INTERVAL_D1, n + 2)),
                          ("Hyperliquid", lambda: _hl_d1(coin, n + 2))):
        try:
            cs = fetch()
        except Exception as e:
            errors.append(f"{source} : {e}")
            continue
        if len(cs) >= n + 1 and cs[-1]["ts"] == yesterday:
            break
    else:
        if len(errors) == 2:
            log_error(f"tendance {coin} : bougies D1 indisponibles ({' ; '.join(errors)})")
        return None
    votes = sum(1 if cs[-1]["c"] > cs[-1 - lb]["c"] else -1 for lb in TREND_LOOKBACKS)
    side = "BUY" if votes > 0 else "SELL"
    _trend_sig[coin] = (day, side, cs[-1]["c"], _atr(cs, 14))
    return _trend_sig[coin][1:]

def _worst_px(px: float, is_buy: bool, slip: float) -> float:
    """Prix le plus bas auquel un ordre au marché (prix limite à `slip`) peut s'exécuter : sa valeur
    minimale vaut qty × ce prix. Une vente peut descendre jusqu'à sa limite ; un achat, pas sous px."""
    return px if is_buy else px * (1 - slip)

def _ceil_qty(coin: str, qty: float) -> float:
    """Quantité arrondie au pas du coin, vers le haut."""
    dec = _flt(coin)["sz_dec"]
    return round(math.ceil(round(qty * 10 ** dec, 6)) / 10 ** dec, dec)

def _trend_qty(coin: str, price: float, equity: float, floor_px: float | None = None) -> float:
    """Quantité pour TREND_EXPOSURE × équité / nb de coins, arrondie au pas du coin vers le haut,
    et valant au moins TREND_MIN_TARGET même au prix floor_px (pire prix de l'ouverture et du SL).
    0 si l'équité ne permet pas ce minimum."""
    floor_px = min(price, floor_px or price)
    if price <= 0 or floor_px <= 0:
        return 0.0
    qty = max(TREND_EXPOSURE * equity / len(TREND_COINS) / price, TREND_MIN_TARGET / floor_px)
    if qty * price > equity / len(TREND_COINS):
        return 0.0
    return _ceil_qty(coin, qty)

def _trend_alert_once(coin: str, why: str, title: str, desc: str, color: int = 0xf39c12):
    key = (coin, _today(), why)
    if key not in _trend_alerted:
        _trend_alerted.add(key)
        send_discord_alert(title, desc, color=color)

def _place_sl(coin: str, is_buy: bool, qty: float, sl: float):
    """Pose le stop-loss (trigger reduce-only) d'une position ; 2 essais. Retourne l'oid, ou None."""
    sl_lim = round_px(sl * (1 - TREND_SLIP_SL if is_buy else 1 + TREND_SLIP_SL), coin)
    for attempt in range(2):
        try:
            r = exchange.order(coin, not is_buy, qty, limit_px=sl_lim,
                               order_type={"trigger": {"triggerPx": float(sl), "isMarket": True, "tpsl": "sl"}},
                               reduce_only=True)
            st = (r.get("response") or {}).get("data", {}).get("statuses", [{}])
            if r.get("status") == "ok" and "error" not in st[0]:
                return (st[0].get("resting") or st[0].get("filled") or {}).get("oid")
            err = r
        except Exception as e:
            err = e
        log_error(f"SL tendance {coin} refusé (essai {attempt + 1}) : {err}")
        if attempt == 0:
            time.sleep(1)
    return None

def trend_open(wb: Workbook, coin: str, side: str, atr: float, equity: float, mids: dict) -> dict:
    """Ouvre la position tendance (fantôme si GHOST_ONLY) avec son SL. Lève une exception si
    l'ouverture échoue ou si le SL est refusé (la position est alors refermée au marché). Ordre
    resté sans réponse : il reste noté dans « opening », trend_cycle le vérifie au cycle suivant."""
    is_buy, sgn = side == "BUY", (1 if side == "BUY" else -1)
    price = float(mids.get(coin) or 0)
    sl_px = price - sgn * TREND_SL_ATR * atr
    # L'ouverture et le SL (ordre inverse) valent chacun ≥ TREND_MIN_NOTIONAL à leur pire prix :
    # long → SL vendu jusqu'à 10 % sous son déclenchement ; short → vente jusqu'à 2 % sous le prix
    floor = min(_worst_px(price, is_buy, TREND_SLIP_OPEN), _worst_px(sl_px, not is_buy, TREND_SLIP_SL))
    qty   = _trend_qty(coin, price, equity, floor)
    if not qty:
        raise RuntimeError(f"équité ${equity:.2f} insuffisante pour {TREND_MIN_NOTIONAL:g} $ par ordre")
    now = time.time()
    if GHOST_ONLY:
        oid, entry, open_fee = f"P{int(now * 1000)}-{coin}", price, TREND_FEE * qty * price
    else:
        set_leverage(coin)
        # Noté AVANT l'ordre : réponse perdue (délai dépassé, 5xx) ou bot arrêté pendant l'ordre →
        # la position éventuellement ouverte est reprise ensuite (avant : prise pour celle d'un autre
        # bot, jamais gérée, sans SL)
        opening = _load_trend()["opening"]
        opening[coin] = {"side": side, "atr": atr, "ts": now}
        _save_trend()
        r  = exchange.market_open(coin, is_buy, qty, px=None, slippage=TREND_SLIP_OPEN)
        del opening[coin]   # réponse reçue ; enregistré par l'appelant avec la position ou l'échec
        st = ((r or {}).get("response") or {}).get("data", {}).get("statuses", [{}])
        if not r or r.get("status") != "ok" or "error" in st[0]:
            raise RuntimeError(f"ordre refusé : {r}")
        fill     = st[0].get("filled") or {}
        oid      = str(fill.get("oid") or f"hl-{coin}-{int(now)}")
        entry    = float(fill.get("avgPx") or price)
        qty      = float(fill.get("totalSz") or qty)
        open_fee = float(fill.get("fee") or TREND_FEE * qty * entry)   # réponse d'ordre sans frais : estimés
    sl  = round_px(entry - sgn * TREND_SL_ATR * atr, coin)
    pos = {"id": oid, "symbol": coin, "side": side, "entry": entry, "qty": qty, "sl": sl,
           "sl_oid": None, "atr": atr, "open_fee": open_fee, "opened_at": now, "paper": GHOST_ONLY}
    if not GHOST_ONLY:
        pos["sl_oid"] = _place_sl(coin, is_buy, qty, sl)
        if pos["sl_oid"] is None:
            # Position sans stop : on la referme tout de suite
            try:
                rc = exchange.market_close(coin, slippage=TREND_SLIP_CLOSE)
                closed = rc is not None and rc.get("status") == "ok"
            except Exception as e:
                log_error(f"market_close urgence {coin}: {e}")
                closed = False
            if closed:
                raise RuntimeError("SL refusé 2× par Hyperliquid, position refermée au marché")
            send_discord_alert(f"🚨 {coin} OUVERT SANS STOP-LOSS",
                               "SL tendance refusé ET fermeture d'urgence échouée. "
                               "**Intervention manuelle requise.** (nouvel essai de SL toutes les 5 min)",
                               color=0xff0000)
    _trend_log_open(wb, pos, equity)
    return pos

def _trend_log_open(wb: Workbook, pos: dict, equity: float, note: str = ""):
    """Ligne Excel, console et Discord d'une position tendance ouverte (`note` : ligne de plus sur Discord)."""
    coin, side, entry, qty, sl = pos["symbol"], pos["side"], pos["entry"], pos["qty"], pos["sl"]
    is_buy, paper = side == "BUY", pos.get("paper")
    notional = round(qty * entry, 2)
    try:
        wb["Tendance"].append([pos["id"], datetime.fromtimestamp(pos["opened_at"], timezone.utc).strftime("%Y-%m-%d %H:%M"),
                               f"{coin}/USDT", side, "Fantôme" if paper else "Réel", qty,
                               round(entry, 6), sl, notional, "Ouvert", None, None, None, None])
        _safe_save(wb)
    except Exception as e:
        log_error(f"excel tendance {coin}: {e}")
    _sc = C.BGRN if is_buy else C.BRED
    print(f"  {_sc}[TENDANCE {'LONG' if is_buy else 'SHORT'}]{C.RST} {C.BOLD}{coin}{C.RST} {qty} @ ${entry:.6g}"
          f"  SL ${sl:.6g}  notionnel ${notional}" + ("  (fantôme)" if paper else ""))
    risk = qty * abs(entry - sl) + TREND_FEE * qty * (entry + sl)   # perte si le SL est touché, frais compris
    send_discord_alert(
        f"{'👻 ' if paper else ''}{'🟢 LONG' if is_buy else '🔴 SHORT'} {coin} ouvert [tendance]",
        f"Entrée : **${entry:.6g}** · SL : **${sl:.6g}** ({abs(sl - entry) / entry * 100:.1f} %)\n"
        f"Notionnel : **${notional}** ({notional / equity * 100:.1f} % du capital) · "
        f"perte au SL : **{-risk / equity * 100:.2f} %** du capital"
        + (" · fantôme (aucun ordre réel)" if paper else "") + note,
        color=0x2ecc71 if is_buy else 0xe74c3c)

def _trend_adopt(wb: Workbook, coin: str, o: dict, p: dict, equity: float) -> dict:
    """Position ouverte par un ordre resté sans réponse (`o`, noté dans « opening ») : reprise telle
    qu'elle est sur le compte (`p`), avec son SL, posé s'il ne l'a pas encore été."""
    global _trend_sl_ts
    sgn = 1 if o["side"] == "BUY" else -1
    entry, qty = float(p["entryPx"]), abs(float(p["szi"]))
    pos = {"id": f"hl-{coin}-{int(o['ts'])}", "symbol": coin, "side": o["side"], "entry": entry, "qty": qty,
           "sl": round_px(entry - sgn * TREND_SL_ATR * o["atr"], coin), "sl_oid": None, "atr": o["atr"],
           "open_fee": TREND_FEE * qty * entry, "opened_at": o["ts"], "paper": False}
    _trend_log_open(wb, pos, equity, "\nHyperliquid n'avait pas répondu à l'ordre : position retrouvée sur le compte")
    ensure_sl_orders([pos])
    if pos["sl_oid"] is None:
        _trend_sl_ts = 0.0   # SL pas encore posé : nouvel essai dès ce cycle, puis toutes les 5 min
    return pos

def _trend_rescale_paper(wb: Workbook, positions: dict, equity: float) -> list[str]:
    """Positions fantômes recalculées pour le capital `equity` (capital fictif du test changé) : même
    entrée, même SL ; quantité et frais d'ouverture refaits comme à l'ouverture. Retourne les coins."""
    done = []
    for coin, pos in sorted(positions.items()):
        if not pos.get("paper"):
            continue
        is_buy = pos["side"] == "BUY"
        floor  = min(_worst_px(pos["entry"], is_buy, TREND_SLIP_OPEN), _worst_px(pos["sl"], not is_buy, TREND_SLIP_SL))
        qty    = _trend_qty(coin, pos["entry"], equity, floor)
        if not qty:
            continue   # capital trop petit pour le minimum par ordre : taille gardée
        pos["qty"], pos["open_fee"] = qty, TREND_FEE * qty * pos["entry"]
        done.append(coin)
        for row in wb["Tendance"].iter_rows(min_row=2):
            if str(row[0].value) == str(pos["id"]):
                row[5].value, row[8].value = qty, round(qty * pos["entry"], 2)
                break
    try:
        _safe_save(wb)
    except Exception as e:
        log_error(f"excel tendance (capital du test) : {e}")
    return done

def _trend_funding(pos: dict) -> float | None:
    """Financement reçu (+) ou payé (−) par la position depuis son ouverture : paiements du compte
    (réelle ; à un jour près, Hyperliquid regroupant les anciens paiements par jour à 00:00 UTC) ou
    taux horaires publics × notionnel (fantôme). None si Hyperliquid ne répond pas."""
    coin, start, total = pos["symbol"], int((pos.get("opened_at") or 0) * 1000), 0.0
    sgn = 1 if pos["side"] == "BUY" else -1
    try:
        for _ in range(50):          # pages de 500 lignes au plus (taux horaires : ≈ 3 semaines)
            if pos.get("paper"):
                page   = info.funding_history(coin, start)
                total -= sgn * pos["qty"] * pos["entry"] * sum(float(f["fundingRate"]) for f in page)
            else:
                page   = info.user_funding_history(WALLET_ADDRESS, start)
                total += sum(float(f["delta"]["usdc"]) for f in page if f["delta"].get("coin") == coin)
            if len(page) < 500:
                return total
            start = int(page[-1]["time"]) + 1
    except Exception as e:
        log_error(f"financement tendance {coin} (non compté) : {e}")
    return None

def _trend_record_close(wb: Workbook, pos: dict, status: str, exit_px: float, pnl: float,
                        equity: float) -> dict:
    """Ligne Excel, console et Discord d'une position tendance fermée (P&L aussi en % de l'équité),
    financement compris (le backtest le compte aussi)."""
    coin, now = pos["symbol"], time.time()
    funding   = _trend_funding(pos) or 0.0
    pnl      += funding
    notional  = pos["entry"] * pos["qty"]
    pnl_pct   = pnl / notional * 100 if notional else 0.0
    base      = (_load_trend().get("paper_equity") or equity) if pos.get("paper") else equity   # fantôme : capital fictif
    cap_pct   = pnl / base * 100 if base > 0 else 0.0
    pos.update(status=status, exit=exit_px, pnl=round(pnl, 4), funding=round(funding, 4), closed_at=now)
    try:
        for row in wb["Tendance"].iter_rows(min_row=2):
            if str(row[0].value) == str(pos["id"]):
                row[9].value, row[10].value = status, round(exit_px, 6) if exit_px else None
                row[11].value, row[12].value = round(pnl, 2), round(pnl_pct, 2)
                row[13].value = datetime.fromtimestamp(now, timezone.utc).strftime("%Y-%m-%d %H:%M")
                for cell in row:
                    cell.fill = _FILLS["tp"] if pnl > 0 else _FILLS["sl"]
                break
        _safe_save(wb)
    except Exception as e:
        log_error(f"excel tendance fermeture {coin}: {e}")
    days = (now - (pos.get("opened_at") or now)) / 86400
    print(f"  {_cp(pnl)}[TENDANCE {status}]{C.RST} {C.BOLD}{coin}{C.RST}  P&L={_cp(pnl)}${pnl:+.2f}{C.RST} "
          f"({pnl_pct:+.2f} % du notionnel · {cap_pct:+.2f} % du capital) · tenu {days:.1f} j")
    send_discord_alert(
        f"{'👻 ' if pos.get('paper') else ''}{'✅' if pnl >= 0 else '🛑'} {coin} {pos['side']} — {status} [tendance]",
        f"P&L : **{cap_pct:+.2f} %** du capital · {pnl_pct:+.2f} % du notionnel (${pnl:+.2f}) · tenu {days:.1f} j\n"
        f"Entrée : ${pos['entry']:.6g}" + (f" → sortie ${exit_px:.6g}" if exit_px else "")
        + (f"\nFinancement {'reçu' if funding > 0 else 'payé'} : ${abs(funding):.2f} (compris dans le P&L)"
           if round(funding, 2) else "")
        + (f"\nRenforcée de {pos['top_up']:g} juste avant : la fermeture seule valait moins de "
           f"{TREND_MIN_NOTIONAL:g} $ (+1 % de marge)" if pos.get("top_up") else ""),
        color=0x2ecc71 if pnl >= 0 else 0xe74c3c)
    return pos

def _paper_pnl(pos: dict, exit_px: float) -> float:
    sgn = 1 if pos["side"] == "BUY" else -1
    return sgn * pos["qty"] * (exit_px - pos["entry"]) - TREND_FEE * pos["qty"] * exit_px - pos.get("open_fee", 0.0)

def _trend_fill_result(pos: dict, fills: list) -> tuple[float, float, float] | None:
    """(P&L net de frais, prix moyen de sortie, heure de sortie en s) de la dernière fermeture de
    la position d'après les fills du compte ; None si elle n'y figure pas (encore)."""
    res, order = _closing_pnl(pos, fills), _closing_fills(pos, fills)
    if not res:
        return None
    gross, open_fee, close_fee = res
    sz = sum(float(f.get("sz") or 0) for f in order)
    px = sum(float(f.get("px") or 0) * float(f.get("sz") or 0) for f in order) / sz if sz else 0.0
    return gross - open_fee - close_fee - pos.get("top_up_fee", 0.0), px, max(f["time"] for f in order) / 1000

def _trend_top_up(pos: dict, px: float):
    """Si l'ordre de fermeture vaudrait moins de TREND_MIN_TARGET (short qui a beaucoup gagné : sa
    valeur a fondu), renforce d'abord la position d'un ordre qui vaut lui-même ce minimum ; la
    fermeture emporte ensuite le tout. Aucun ordre sous le minimum, pour ~0.01 $ de frais en plus."""
    coin, is_buy = pos["symbol"], pos["side"] == "BUY"
    held = pos["qty"] + pos.get("top_up", 0.0)   # renfort d'un essai précédent dont la fermeture a échoué
    if not px or held * _worst_px(px, not is_buy, TREND_SLIP_CLOSE) >= TREND_MIN_TARGET:
        return
    add = _ceil_qty(coin, TREND_MIN_TARGET / min(_worst_px(px, is_buy, TREND_SLIP_OPEN),
                                                 _worst_px(px, not is_buy, TREND_SLIP_CLOSE)))
    r  = exchange.market_open(coin, is_buy, add, px=None, slippage=TREND_SLIP_OPEN)
    st = ((r or {}).get("response") or {}).get("data", {}).get("statuses", [{}])
    if not r or r.get("status") != "ok" or "error" in st[0]:
        raise RuntimeError(f"renfort avant fermeture refusé : {r}")
    add = float((st[0].get("filled") or {}).get("totalSz") or add)
    pos["top_up"]     = pos.get("top_up", 0.0) + add
    pos["top_up_fee"] = pos.get("top_up_fee", 0.0) + TREND_FEE * add * px
    print(f"  {C.BYLW}[TENDANCE RENFORT]{C.RST} {coin} +{add} avant fermeture (ordre sous {TREND_MIN_NOTIONAL:g} $ + 1 % évité)")

def _trend_cancel_sl(coin: str, oid):
    """Annule le SL d'une position fermée. Échec (réseau) : nouvel essai à chaque cycle (« cancels ») ;
    resté en place, ce SL reduce-only fermerait plus tard une position du même sens à l'ancien niveau."""
    if not oid:
        return
    try:
        exchange.cancel(coin, int(oid))
    except Exception as e:
        log_error(f"annulation SL tendance {coin} (nouvel essai au cycle suivant) : {e}")
        _load_trend()["cancels"].append([coin, oid])

def trend_close(wb: Workbook, pos: dict, status: str, mids: dict, equity: float) -> dict:
    """Ferme la position tendance au marché (fantôme : au prix moyen), après un renfort si l'ordre de
    fermeture vaudrait moins de TREND_MIN_TARGET. Lève une exception si Hyperliquid refuse
    (la position et son SL restent en place)."""
    coin    = pos["symbol"]
    exit_px = float(mids.get(coin) or 0)
    if pos.get("paper"):
        return _trend_record_close(wb, pos, status, exit_px, _paper_pnl(pos, exit_px or pos["entry"]), equity)
    _trend_top_up(pos, exit_px)
    # Noté AVANT l'ordre : réponse perdue (délai dépassé, 5xx) → au cycle suivant, une position absente
    # du compte a été fermée par cet ordre, pas par son SL (avant : « SL touché », à plat jusqu'à demain)
    pos["closing"] = status
    _save_trend()
    r  = exchange.market_close(coin, slippage=TREND_SLIP_CLOSE)
    st = ((r or {}).get("response") or {}).get("data", {}).get("statuses", [{}])
    if r is not None and (r.get("status") != "ok" or (st and "error" in st[0])):
        del pos["closing"]   # refus net : la position est toujours là
        _save_trend()
        raise RuntimeError(f"fermeture refusée : {r}")
    _trend_cancel_sl(coin, pos.get("sl_oid"))
    time.sleep(1)   # propagation des fills
    res = None
    try:
        res = _trend_fill_result(pos, info.user_fills(WALLET_ADDRESS))
    except Exception as e:
        log_error(f"P&L tendance {coin}: {e}")
    if res is None:   # fills pas encore visibles : P&L estimé au prix moyen (avant : 0 $)
        log_error(f"tendance {coin} : fermeture absente des fills, P&L estimé au prix moyen")
        res = (_paper_pnl(pos, exit_px or pos["entry"]) - pos.get("top_up_fee", 0.0), exit_px, 0.0)
    pnl, px, _ = res
    return _trend_record_close(wb, pos, status, px or exit_px, pnl, equity)

def _trend_try_close(wb: Workbook, pos: dict, status: str, mids: dict, equity: float, closed: list) -> bool:
    """trend_close ; si Hyperliquid refuse, une alerte par jour et False (la position garde son
    SL, nouvel essai au cycle suivant)."""
    try:
        closed.append(trend_close(wb, pos, status, mids, equity))
        return True
    except Exception as e:
        log_error(f"fermeture tendance {pos['symbol']}: {e}")
        if pos.get("closing"):   # ordre peut-être passé
            _trend_alert_once(pos["symbol"], "close", f"⚠️ Tendance : fermeture {pos['symbol']} sans réponse",
                              f"Hyperliquid n'a pas répondu : la fermeture est vérifiée au cycle suivant "
                              f"(la position garde son SL d'ici là).\nErreur : `{str(e)[:300]}`")
        else:
            _trend_alert_once(pos["symbol"], "close", f"⚠️ Tendance : fermeture {pos['symbol']} échouée",
                              f"La position reste ouverte avec son SL ; nouvel essai à chaque cycle.\n"
                              f"Erreur : `{str(e)[:300]}`")
        return False

def trend_cycle(wb: Workbook, user_state: dict | None, equity: float) -> list[dict]:
    """Un passage de la stratégie tendance (chaque cycle) : détecte les SL touchés, puis aligne
    chaque coin sur le signal du jour. Retourne les positions fermées pendant ce passage."""
    global _trend_sl_ts
    state, day = _load_trend(), _today()
    positions  = state["positions"]
    state["stopped"] = {c: d for c, d in state["stopped"].items() if d == day}
    closed: list[dict] = []
    acct = {ap["position"]["coin"]: ap["position"] for ap in (user_state or {}).get("assetPositions", [])
            if float(ap["position"].get("szi") or 0)}   # positions du compte (de tous les bots du wallet)
    held_on_account = set(acct) if user_state is not None else None
    try:
        mids = info.all_mids()
    except Exception as e:
        log_error(f"tendance all_mids (nouvel essai au cycle suivant) : {e}")
        return closed

    # Capital fictif du test changé (mode fantôme) : positions fantômes recalculées à cette taille
    if GHOST_ONLY and state.get("paper_equity", 0) != PAPER_EQUITY:
        coins = _trend_rescale_paper(wb, positions, equity)
        state["paper_equity"] = PAPER_EQUITY
        _save_trend()
        send_discord_alert(
            f"🧪 Test : capital fictif de ${PAPER_EQUITY:g}" if PAPER_EQUITY > 0 else "🧪 Test : retour au solde réel",
            (f"Le mode fantôme calcule ses positions comme si le compte avait **${PAPER_EQUITY:g}** (aucun ordre "
             f"réel ; ignoré au passage en réel)." if PAPER_EQUITY > 0 else
             f"Le mode fantôme calcule de nouveau ses positions sur le solde réel (${equity:.2f}).")
            + (f"\nPositions fantômes en cours recalculées à cette taille (même entrée, même SL) : {', '.join(coins)}"
               if coins else ""),
            color=0x9b59b6)

    # 0) Ordres restés sans réponse (délai dépassé, bot arrêté pendant l'ordre). Ouverture : position
    #    reprise si elle est sur le compte, sinon nouvel essai plus bas. SL à annuler : nouvel essai.
    if state["cancels"]:
        try:   # seulement les SL encore en place (sinon déjà annulés ou exécutés : rien à refaire)
            resting = {o.get("oid") for o in info.frontend_open_orders(WALLET_ADDRESS)}
        except Exception as e:
            log_error(f"ordres ouverts (annulation de SL au cycle suivant) : {e}")
            resting = None
        for coin, oid in list(state["cancels"]):
            if resting is not None:
                state["cancels"].remove([coin, oid])
                if oid in resting:
                    _trend_cancel_sl(coin, oid)
        _save_trend()
    for coin, o in list(state["opening"].items() if held_on_account is not None else ()):
        p = acct.get(coin)
        if coin not in positions and p and (float(p["szi"]) > 0) == (o["side"] == "BUY"):
            positions[coin] = _trend_adopt(wb, coin, o, p, equity)
        del state["opening"][coin]
        _save_trend()

    # 1) SL touchés (réel : position absente du compte ; fantôme : prix moyen au-delà du SL)
    for coin, pos in list(positions.items()):
        if pos.get("paper"):
            px = float(mids.get(coin) or 0)
            if px and (px <= pos["sl"] if pos["side"] == "BUY" else px >= pos["sl"]):
                closed.append(_trend_record_close(wb, pos, "SL touché", pos["sl"], _paper_pnl(pos, pos["sl"]), equity))
                del positions[coin]
                state["stopped"][coin] = day
            continue
        if held_on_account is None or coin in held_on_account:
            pos.pop("_missing", None)
            if held_on_account is not None:
                pos.pop("closing", None)   # fermeture sans réponse pas passée : nouvel essai plus bas
            # Renfort passé sans réponse : la taille du compte fait foi (sinon second renfort inutile)
            extra = round(abs(float(acct[coin]["szi"])) - pos["qty"], 9) if coin in acct else 0.0
            if extra > pos.get("top_up", 0.0) + 1e-9:
                pos["top_up_fee"] = (pos.get("top_up_fee", 0.0)
                                     + TREND_FEE * (extra - pos.get("top_up", 0.0)) * float(mids.get(coin) or pos["entry"]))
                pos["top_up"] = extra
            continue
        res = None
        try:
            res = _trend_fill_result(pos, info.user_fills(WALLET_ADDRESS))
        except Exception as e:
            log_error(f"fills tendance {coin}: {e}")
        why = pos.get("closing")   # fermée par notre ordre resté sans réponse (motif de cette fermeture)
        if not res:
            # Fills pas encore propagés ou erreur API passagère : on attend 3 cycles
            pos["_missing"] = pos.get("_missing", 0) + 1
            if pos["_missing"] < 3:
                continue
            px = (float(mids.get(coin) or 0) if why else 0.0) or pos["sl"]
            log_error(f"tendance {coin} : absente du compte sans fill de fermeture, P&L estimé à ${px:.6g}")
            res = (_paper_pnl(pos, px) - pos.get("top_up_fee", 0.0), px, time.time())
        pnl, exit_px, exit_ts = res
        closed.append(_trend_record_close(wb, pos, why or ("SL touché" if pnl < 0 else "Fermé (SL)"),
                                          exit_px, pnl, equity))
        del positions[coin]
        # Fermée par autre chose que son SL (fermeture restée sans réponse, à la main) : SL orphelin annulé.
        # Après un vrai SL, Hyperliquid répond simplement « déjà exécuté ».
        _trend_cancel_sl(coin, pos.get("sl_oid"))
        # Après un SL : à plat jusqu'à demain, sauf SL d'un jour passé (bot arrêté à ce moment). Après
        # notre fermeture : aucune attente, la nouvelle position s'ouvre plus bas dans ce cycle.
        if not why and _utc_day(exit_ts) == day:
            state["stopped"][coin] = day
        _save_trend()

    # Surveillance des SL réels toutes les 5 min (position sans stop = risque max)
    real = [p for p in positions.values() if not p.get("paper")]
    if real and time.time() - _trend_sl_ts >= SL_CHECK_S:
        _trend_sl_ts = time.time()
        ensure_sl_orders(real)

    # Coins retirés de TREND_COINS (LTC le 2026-11-03) : fermés, jamais rouverts
    for coin in [c for c in positions if c not in TREND_COINS]:
        if _trend_try_close(wb, positions[coin], "retiré de la stratégie", mids, equity, closed):
            del positions[coin]
            _save_trend()

    # 2) Signal du jour. État enregistré après chaque ordre : une position ouverte mais absente de
    #    trend_state.json (bot arrêté juste après l'ordre) serait ensuite prise pour celle d'un autre bot.
    # Tailles des coins non chargées au démarrage (Hyperliquid injoignable) : nouvel essai à chaque
    # cycle, et aucune ouverture réelle d'ici là (quantité au mauvais pas : ordre refusé ou trop gros).
    sizes_ok = _sym_filters_ok or init_exchange_info()
    just_closed: set[str] = set()
    for coin in TREND_COINS:
        sig = trend_signal(coin)
        if not sig and datetime.now(timezone.utc).hour >= 1:   # toujours rien une heure après minuit
            _trend_alert_once(coin, "signal", f"⚠️ Tendance : pas de signal {coin}",
                              "Bougies journalières d'hier indisponibles (Bybit et Hyperliquid) : aucune "
                              "décision sur ce coin, sa position éventuelle garde son SL. Nouvel essai à chaque cycle.")
        if not sig or not float(mids.get(coin) or 0):
            continue   # bougies ou prix indisponibles : nouvel essai au cycle suivant
        side, _close, atr = sig
        pos = positions.get(coin)
        if pos and pos["side"] == side and not (pos.get("paper") and not GHOST_ONLY):
            continue
        if pos:
            status = "signal inversé" if pos["side"] != side else "fin du mode fantôme"
            if not _trend_try_close(wb, pos, status, mids, equity, closed):
                continue
            del positions[coin]
            _save_trend()
            if not pos.get("paper"):
                just_closed.add(coin)   # encore visible sur le compte lu en début de cycle
        if state["stopped"].get(coin) == day:
            continue
        if not GHOST_ONLY and (held_on_account is None or not sizes_ok or
                               (coin in held_on_account and coin not in just_closed)):
            continue   # compte illisible, tailles inconnues, ou position d'un autre bot (BotCopyTrade)
        try:
            positions[coin] = trend_open(wb, coin, side, atr, equity, mids)
        except Exception as e:
            unsure = coin in state["opening"]   # ordre sans réponse : vérifié au cycle suivant
            if not unsure:
                state["stopped"][coin] = day   # refus : pas de nouvel essai avant demain (évite d'insister)
            log_error(f"ouverture tendance {coin}: {e}")
            _trend_alert_once(coin, "open", f"⚠️ Tendance : {coin} non ouvert",
                              f"{side} {coin} : `{str(e)[:300]}`\n"
                              + ("Hyperliquid n'a pas répondu : au cycle suivant, la position est reprise avec "
                                 "son SL si l'ordre est passé, sinon l'ouverture est retentée." if unsure
                                 else "Nouvel essai demain."))
        _save_trend()
    _save_trend()
    return closed

def trend_summary(mids: dict | None) -> str:
    """« ETH LONG +1.2 % · SOL SHORT −0.4 % » (variation depuis l'entrée, dans le sens de la position)."""
    parts = []
    for coin, p in sorted(_load_trend()["positions"].items()):
        px  = float((mids or {}).get(coin) or 0)
        chg = ((px - p["entry"]) / p["entry"] * 100 * (1 if p["side"] == "BUY" else -1)) if px else 0.0
        parts.append(f"{coin} {'LONG' if p['side'] == 'BUY' else 'SHORT'} {chg:+.1f} %")
    return " · ".join(parts) or "aucune position"

def trend_heartbeat(wb: Workbook, mids: dict | None, equity: float, closed: list[dict]) -> str:
    """Bloc tendance du résumé Discord : variation de chaque position depuis l'entrée, puis latent,
    fermetures de la période et bilan du mode en cours (fantôme / réel) en % de l'équité."""
    pct  = lambda usd: usd / equity * 100 if equity > 0 else 0.0
    poss = _load_trend()["positions"]
    lat  = {c: (1 if p["side"] == "BUY" else -1) * p["qty"] * (float(mids[c]) - p["entry"])
            for c, p in poss.items() if (mids or {}).get(c)}
    mode = "Fantôme" if GHOST_ONLY else "Réel"
    done: list[float] = []
    try:
        done = [r[11] for r in wb["Tendance"].iter_rows(min_row=2, values_only=True)
                if r[4] == mode and r[11] is not None]
    except Exception as e:
        log_error(f"bilan tendance (Excel) : {e}")
    wins    = sum(1 for x in done if x > 0)
    mode_lt = sum(v for c, v in lat.items() if bool(poss[c].get("paper")) == GHOST_ONLY)
    period  = sum(p["pnl"] for p in closed)
    return (f"Tendance : {trend_summary(mids)}\n"
            f"Latent : **{pct(sum(lat.values())):+.2f} %** du capital\n"
            f"Fermetures tendance : {len(closed)}"
            + (f" · **{pct(period):+.2f} %** du capital (${period:+.2f})" if closed else "") + "\n"
            f"Bilan {mode.lower()} depuis le début : **{pct(sum(done) + mode_lt):+.2f} %** du capital "
            f"({len(done)} fermée{'s' if len(done) > 1 else ''}, {wins} gagnante{'s' if wins > 1 else ''})")

# ── Tableau de bord (dashboard.html) ──────────────────────────────────────────
def write_dashboard(wb: Workbook, user_state: dict | None, balance: float, initial_balance: float):
    """Page locale régénérée à chaque cycle et rafraîchie toutes les 15 s dans le navigateur : équité,
    positions tendance avec leur P&L latent lu sur le compte, bilan des positions fermées (avant : page de
    l'ancien scalping, régime, scans et leçons). Une erreur est journalisée sans arrêter le bot."""
    try:
        upnl = {ap["position"]["coin"]: float(ap["position"].get("unrealizedPnl") or 0)
                for ap in (user_state or {}).get("assetPositions", [])}
        mode = "Fantôme" if GHOST_ONLY else "Réel"
        done = [r[11] for r in wb["Tendance"].iter_rows(min_row=2, values_only=True)
                if r[4] == mode and r[11] is not None]
        pct  = lambda usd: usd / balance * 100 if balance > 0 else 0.0
        col  = lambda v: "#4ade80" if v > 0 else ("#f87171" if v < 0 else "#e5e7eb")
        rows = ""
        for coin, p in sorted(_load_trend()["positions"].items()):
            lat = None if p.get("paper") else upnl.get(coin)
            rows += (f"<tr><td>{coin}</td><td>{'LONG' if p['side'] == 'BUY' else 'SHORT'}</td><td>{p['qty']:g}</td>"
                     f"<td>${p['entry']:.6g}</td><td>${p['sl']:.6g}</td>"
                     + (f"<td style='color:{col(lat)}'>${lat:+.2f} ({pct(lat):+.2f} %)</td>" if lat is not None
                        else "<td>—</td>")
                     + f"<td>{datetime.fromtimestamp(p.get('opened_at') or 0, timezone.utc):%d/%m %H:%M}</td></tr>")
        html = f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><meta http-equiv="refresh" content="15"><title>BotShortTrade</title>
<style>body{{background:#111827;color:#e5e7eb;font-family:system-ui,sans-serif;margin:24px}}
td,th{{padding:6px 14px;border-bottom:1px solid #374151;text-align:left}}table{{border-collapse:collapse}}
.k{{display:inline-block;margin:16px 40px 16px 0}}.k b{{display:block;font-size:26px}}small{{color:#9ca3af}}</style></head>
<body><h1>BotShortTrade — tendance journalière</h1>
<small>Mis à jour le {datetime.now(timezone.utc):%d/%m/%Y à %H:%M:%S} UTC · rafraîchi toutes les 15 s ·
{'TESTNET' if HL_TESTNET else 'MAINNET (argent réel)'}{' · mode fantôme (aucun ordre réel)' if GHOST_ONLY else ''}</small><br>
<div class="k">Équité<b>${balance:.2f}</b><small>{_since_start(balance, initial_balance)} (réf. ${initial_balance:.2f})</small></div>
<div class="k">Positions fermées ({mode.lower()})<b style="color:{col(sum(done))}">${sum(done):+.2f}</b>
<small>{len(done)} fermée(s) · {sum(1 for x in done if x > 0)} gagnante(s)</small></div>
<table><tr><th>Coin</th><th>Sens</th><th>Qté</th><th>Entrée</th><th>SL</th><th>Latent</th><th>Ouverte (UTC)</th></tr>
{rows or "<tr><td colspan='7'>aucune position</td></tr>"}</table>
<p><small>Vote {'/'.join(map(str, TREND_LOOKBACKS))} j · SL {TREND_SL_ATR:g}×ATR D1 · exposition {TREND_EXPOSURE:.0%} de l'équité</small></p>
</body></html>"""
        with open("dashboard.html", "w", encoding="utf-8") as f:
            f.write(html)
    except Exception as e:
        log_error(f"write_dashboard: {e}")

# ── Boucle principale ─────────────────────────────────────────────────────────
def run():
    print(f"\n{C.BCYN}{'═' * 60}{C.RST}")
    print(f"  {C.BOLD}BotShortTrade  –  Hyperliquid DEX{C.RST}")
    print(f"{C.BCYN}{'═' * 60}{C.RST}")

    wb = init_excel()
    init_exchange_info()

    # Aucune décision sans l'équité réelle : Hyperliquid injoignable au démarrage → nouvel essai (avant :
    # 100 $ supposés, et des positions taillées sur ce montant si la lecture suivante échouait aussi)
    while True:
        try:
            balance = get_equity()
            break
        except Exception as e:
            log_error(f"équité au démarrage (nouvel essai dans {STARTUP_RETRY_S} s) : {e}")
            time.sleep(STARTUP_RETRY_S)
    print(f"  Équité : {C.BOLD}${balance:.2f} USDC{C.RST}"
          + (f" · tailles calculées sur au plus ${_MAX_LIQ_CFG:.2f} (MAX_LIQUIDITY)" if _MAX_LIQ_CFG > 0 else ""))
    initial_balance = load_initial_balance(balance)
    initial_balance += sum(f for _, f in sync_deposits())   # dépôts / retraits faits bot arrêté
    run._ledger_ts = time.time()
    print(f"  {C.DIM}Solde initial (référence % P&L) : ${initial_balance:.2f} USDC{C.RST}")
    print(f"\n  {C.DIM}Démarrage — un cycle toutes les {LOOP_SECONDS}s{C.RST}\n")

    send_discord_alert(
        "🚀 Bot démarré",
        f"Réseau : **{'TESTNET' if HL_TESTNET else 'MAINNET'}**"
        + (" · **mode fantôme seul (aucun ordre réel)**" if GHOST_ONLY else "")
        + (f" · capital fictif du test : **${PAPER_EQUITY:g}**" if GHOST_ONLY and PAPER_EQUITY > 0 else "") + "\n"
        f"Stratégie : **tendance journalière** (vote {'/'.join(map(str, TREND_LOOKBACKS))} j, "
        f"SL {TREND_SL_ATR:g}×ATR D1) · exposition {TREND_EXPOSURE:.0%} de l'équité · positions tendance : "
        f"{len(_load_trend()['positions'])}\n"
        f"Équité : **${balance:.2f}** · {_since_start(balance, initial_balance)}\n"
        f"Coins : {', '.join(TREND_COINS)}",
        color=0x3498db)

    hb_last, hb_cycles = time.time(), 0   # résumé Discord : remis à zéro après chaque envoi
    hb_closed: list[dict] = []            # positions tendance fermées sur la période
    cycle_errors: set[tuple] = set()      # (jour UTC, erreur) imprévues déjà signalées sur Discord

    while True:
        loop_start = time.time()
        try:
            ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
            print(f"\n{C.DIM}[{ts}] ──────────────────────────────────────────────────{C.RST}")

            # Compte lu une fois par cycle : équité, positions (SL touchés, coins tenus par un autre bot)
            user_state: dict | None = None
            try:
                user_state = info.user_state(WALLET_ADDRESS)
            except Exception as e:
                log_error(f"user_state: {e}")
            try:
                balance = get_equity(user_state)
            except Exception as e:
                log_error(f"équité (solde précédent conservé) : {e}")
            print(f"  Équité : {C.BOLD}${balance:.2f}{C.RST} · {_since_start(balance, initial_balance)}")

            # Stratégie tendance : SL touchés, puis alignement de chaque coin sur le signal du jour
            try:
                hb_closed += trend_cycle(wb, user_state, _sizing_equity(balance))
            except Exception as e:
                log_error(f"tendance : {e}")
                traceback.print_exc()
            held = ", ".join(c + " " + p["side"] for c, p in sorted(_load_trend()["positions"].items()))
            print(f"  Tendance{' (fantôme)' if GHOST_ONLY else ''} : {held or 'aucune position'}")

            # Dépôts / retraits relevés toutes les 10 min, et tout de suite si l'équité bouge de plus de 5 % :
            # référence du « % depuis le départ » ajustée (avant : un dépôt comptait comme un gain)
            prev = run.__dict__.get("_prev_equity", 0.0)
            moves: list[tuple[int, float]] = []
            if (prev > 0 and abs(balance - prev) > prev * 0.05) or \
                    time.time() - run.__dict__.get("_ledger_ts", 0) >= LEDGER_REFRESH_S:
                run._ledger_ts = time.time()
                moves = sync_deposits()
                initial_balance += sum(f for _, f in moves)
            # Flash crash : équité en baisse de plus de 5 % en un cycle, hors retrait fait pendant ce cycle
            # (avant : signalé comme un krach)
            if prev > 0:
                since_ms = (run.__dict__.get("_prev_ts", 0) - 60) * 1000
                moved = sum(f for t, f in moves if t >= since_ms)
                equity_drop_pct = (prev - (balance - moved)) / prev * 100
                if equity_drop_pct > 5.0:
                    print(f"  {C.BRED}[🔴 FLASH CRASH]{C.RST} {equity_drop_pct:.1f}% en 1 cycle")
                    send_discord_alert("⚡ Flash Crash Détecté",
                                       f"Équité: **${prev:.2f}** → **${balance:.2f}** "
                                       f"(**-{equity_drop_pct:.1f}%** en 1 cycle)", color=0xe74c3c)
            run._prev_equity, run._prev_ts = balance, time.time()

            # Résumé Discord périodique : positions tendance, fermetures de la période, équité
            hb_cycles += 1
            if HEARTBEAT_HOURS > 0 and time.time() - hb_last >= HEARTBEAT_HOURS * 3600:
                try:
                    mids = info.all_mids()
                except Exception:
                    mids = None
                send_discord_alert(
                    f"💓 Bot actif — résumé {HEARTBEAT_HOURS:g}h",
                    ("Mode : **fantôme seul** (aucun ordre réel)"
                     + (f" · capital fictif **${PAPER_EQUITY:g}**" if PAPER_EQUITY > 0 else "") + "\n"
                     if GHOST_ONLY else "")
                    + trend_heartbeat(wb, mids, _sizing_equity(balance), hb_closed) + "\n"
                    f"Équité : **${balance:.2f}** · {_since_start(balance, initial_balance)}\n"
                    f"Cycles : {hb_cycles}",
                    color=0x95a5a6)
                hb_last, hb_cycles, hb_closed = time.time(), 0, []

            write_dashboard(wb, user_state, balance, initial_balance)
        except Exception as e:
            # Erreur imprévue : journalisée, signalée une fois par jour, nouvel essai au cycle suivant
            # (avant : bot arrêté, et systemd ne le relançait pas, la sortie semblant normale)
            log_error(f"cycle : {e}")
            traceback.print_exc()
            err = f"{type(e).__name__}: {e}"[:300]
            if (_today(), err) not in cycle_errors:
                cycle_errors.add((_today(), err))
                send_discord_alert("⚠️ Erreur inattendue (le bot continue)",
                                   f"`{err}`\nNouvel essai à chaque cycle ; les SL restent en place sur Hyperliquid.",
                                   color=0xf39c12)

        wait = max(0.0, LOOP_SECONDS - (time.time() - loop_start))
        if wait > 0:
            time.sleep(wait)


def main():
    try:
        run()
    except KeyboardInterrupt:
        print(f"\n  {C.BRED}[STOP]{C.RST} Arrêt demandé")
    except Exception as e:
        log_error(f"Crash fatal: {e}")
        traceback.print_exc()
        send_discord_alert("💥 Bot arrêté par une erreur",
                           f"`{type(e).__name__}: {str(e)[:300]}`\nRedémarrage automatique dans 15 s ; "
                           f"les SL restent en place sur Hyperliquid.", color=0xff0000)
        sys.exit(1)   # code d'erreur : systemd (Restart=on-failure) relance le bot (avant : sortie 0, bot arrêté)


if __name__ == "__main__":
    main()
