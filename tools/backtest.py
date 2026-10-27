#!/usr/bin/env python3
"""
Backtest — BotShortTrade
========================
Rejoue une stratégie minute par minute sur l'historique Bybit (données publiques)
et simule les trades sur les bougies M1, frais compris.

Par défaut la stratégie est la VRAIE fonction main.get_signal (mêmes indicateurs,
mêmes seuils) : un résultat ici = ce que le bot aurait fait, hors filtre de régime
(tous les modes sont évalués, le régime réel est noté sur chaque trade).

Usage :
    python3 tools/backtest.py --days 90                       # télécharge (cache) + rapport
    python3 tools/backtest.py --days 30 --sl 4 --tp 4 --fee 0.0576
    python3 tools/backtest.py --strategy h1_trend --max-hours 72
    python3 tools/backtest.py --list                          # stratégies disponibles

Sécurité : main.py est importé comme dans tests/ (SDK Hyperliquid remplacé par des
stubs, clé factice, .env jamais lu, dossier temporaire) → aucun ordre, aucun Discord.
Seul réseau : téléchargement des bougies publiques Bybit (mis en cache dans backtest_data/).

Limites (à garder en tête avant de conclure) :
  - entrée à l'ouverture de la bougie M1 suivant le signal, sans slippage ;
  - SL prioritaire si SL et TP sont touchés dans la même bougie M1 (prudent) ;
  - une position par coin à la fois, pas de limite globale MAX_TRADES ;
  - break-even / trailing simplifiés (déclenchement à 1.5×ATR, trailing à 1×ATR) ;
  - un résultat n'est crédible que si t ≥ 2 ET s'il bat le témoin « entrées au hasard ».
"""
from __future__ import annotations

import argparse, bisect, contextlib, io, json, math, os, pickle, random, statistics
import sys, tempfile, time, types, urllib.request
from unittest import mock

ROOT      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(ROOT, "backtest_data")
COINS     = ["ETH", "LTC", "DOGE", "XRP", "SOL"]
FEE_TAKER = 0.0864   # aller-retour taker Hyperliquid (2 × 0.0432 %)
FEE_MAKER = 0.0576   # entrée maker (0.0144 %) + sortie taker (0.0432 %)
BYBIT     = "https://api.bybit.com/v5/market/kline"
SPAN_MS   = {"1m": 60_000, "5m": 300_000, "15m": 900_000, "1h": 3_600_000}


# ── Import isolé de main.py ──────────────────────────────────────────────────
def load_bot():
    """Importe main.py sans .env, sans SDK réel, dans un dossier temporaire."""
    if "main" in sys.modules:          # déjà importé (ex. par tests/) → réutiliser
        return sys.modules["main"]
    dotenv = types.ModuleType("dotenv"); dotenv.load_dotenv = lambda *a, **k: None
    sys.modules["dotenv"] = dotenv
    os.environ.update({"HL_PRIVATE_KEY": "0x" + "11" * 32, "HL_TESTNET": "1",
                       "DISCORD_WEBHOOK_URL": ""})
    os.environ.setdefault("VOL_MULTIPLIER", "1.2")      # valeur de production
    os.environ.setdefault("MIN_MOVE_PCT", "0.0002")
    class _Stub:
        def __init__(self, *a, **k): pass
    for name, attr in (("hyperliquid.info", "Info"), ("hyperliquid.exchange", "Exchange")):
        mod = types.ModuleType(name); setattr(mod, attr, _Stub); sys.modules[name] = mod
    sys.modules.setdefault("hyperliquid", types.ModuleType("hyperliquid"))
    try:
        import hyperliquid.utils.constants  # noqa: F401  (constantes seules, sans réseau)
    except Exception:
        u = types.ModuleType("hyperliquid.utils"); c = types.ModuleType("hyperliquid.utils.constants")
        c.TESTNET_API_URL = c.MAINNET_API_URL = "http://stub"; u.constants = c
        sys.modules["hyperliquid.utils"] = u; sys.modules["hyperliquid.utils.constants"] = c
    cwd = os.getcwd()
    os.chdir(tempfile.mkdtemp(prefix="bst_backtest_"))
    sys.path.insert(0, ROOT)
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            import main
    finally:
        os.chdir(cwd)
    return main


# ── Données ──────────────────────────────────────────────────────────────────
def _get(url: str) -> dict:
    for attempt in range(3):
        try:
            with urllib.request.urlopen(url, timeout=15) as r:
                d = json.load(r)
            if d.get("retCode") == 0:
                return d
        except Exception:
            pass
        time.sleep(1 + attempt)
    raise RuntimeError(f"Bybit injoignable : {url}")

def download(days: int, refresh: bool = False, cache_dir: str = CACHE_DIR) -> None:
    """Bougies M1 des COINS + BTC (régime), et D1 BTC/ETH. Cache réutilisé s'il couvre
    la période et date de moins de 6 h (sauf refresh)."""
    os.makedirs(cache_dir, exist_ok=True)
    end   = int(time.time() // 60 * 60_000)
    start = end - days * 86_400_000
    for coin in COINS + ["BTC"]:
        path = os.path.join(cache_dir, f"{coin}_1m.json")
        if not refresh and os.path.exists(path) and time.time() - os.path.getmtime(path) < 6 * 3600:
            with open(path) as f:
                first = json.load(f)[0][0]
            if first <= start + 60_000:
                continue
        rows, cur = {}, end
        while cur > start:
            d = _get(f"{BYBIT}?category=linear&symbol={coin}USDT&interval=1&limit=1000&end={cur}")
            lst = d["result"]["list"]
            if not lst:
                break
            for r in lst:
                rows[int(r[0])] = [float(x) for x in r[1:6]]
            cur = min(int(r[0]) for r in lst) - 1
            time.sleep(0.12)
        with open(path, "w") as f:
            json.dump([[k] + rows[k] for k in sorted(rows) if k >= start], f)
        print(f"  {coin} : {len(rows)} bougies M1", file=sys.stderr)
    for coin in ("BTC", "ETH"):
        d = _get(f"{BYBIT}?category=linear&symbol={coin}USDT&interval=D&limit={min(1000, days + 30)}")
        with open(os.path.join(cache_dir, f"{coin}_1d.json"), "w") as f:
            json.dump(sorted([int(r[0]), float(r[4])] for r in d["result"]["list"]), f)

def agg(m1: list[dict], minutes: int) -> list[dict]:
    """Agrège des bougies M1 en bougies de `minutes` (seules les bougies complètes)."""
    out, ms, cur = [], minutes * 60_000, None
    for x in m1:
        b = x["ts"] // ms * ms
        if cur is None or cur["ts"] != b:
            if cur:
                out.append(cur)
            cur = {"ts": b, "o": x["o"], "h": x["h"], "l": x["l"], "c": x["c"], "v": x["v"], "_n": 1}
        else:
            cur["h"] = max(cur["h"], x["h"]); cur["l"] = min(cur["l"], x["l"])
            cur["c"] = x["c"]; cur["v"] += x["v"]; cur["_n"] += 1
    if cur:
        out.append(cur)
    return [c for c in out if c["_n"] == minutes]

def load(days: int | None = None, cache_dir: str = CACHE_DIR) -> tuple[dict, dict]:
    data = {}
    for coin in COINS:
        with open(os.path.join(cache_dir, f"{coin}_1m.json")) as f:
            raw = json.load(f)
        if days:
            raw = [r for r in raw if r[0] >= raw[-1][0] - days * 86_400_000]
        m1 = [{"ts": r[0], "o": r[1], "h": r[2], "l": r[3], "c": r[4], "v": r[5]} for r in raw]
        d = {"1m": m1, "5m": agg(m1, 5), "15m": agg(m1, 15), "1h": agg(m1, 60)}
        for iv in list(d):
            d[iv + "_ts"] = [x["ts"] for x in d[iv]]
        data[coin] = d
    daily = {}
    for coin in ("BTC", "ETH"):
        with open(os.path.join(cache_dir, f"{coin}_1d.json")) as f:
            daily[coin] = [tuple(x) for x in json.load(f)]
    return data, daily

def window(data: dict, coin: str, iv: str, T: int, n: int) -> list[dict]:
    """Les n dernières bougies CLOSES à l'instant T (comme fetch_candles, qui retire
    la bougie en cours)."""
    i = bisect.bisect_right(data[coin][iv + "_ts"], T - SPAN_MS[iv])
    return data[coin][iv][max(0, i - n):i]

def regime_at(daily: dict, T: int) -> str:
    """Même règle que main.get_market_regime, sur les bougies D1 closes avant T."""
    closes, chg = {}, {}
    for c in ("BTC", "ETH"):
        d = [px for ts, px in daily[c] if ts + 86_400_000 <= T][-8:]
        if len(d) < 8:
            return "RANGING"
        closes[c], chg[c] = d, (d[-1] - d[0]) / d[0] * 100
    b, e = chg["BTC"], chg["ETH"]
    if (b > 3 and e < -3) or (b < -3 and e > 3):
        return "DIVERGING"
    avg = (b + e) / 2
    if avg > 5:
        b3 = (closes["BTC"][-1] - closes["BTC"][-4]) / closes["BTC"][-4] * 100
        e3 = (closes["ETH"][-1] - closes["ETH"][-4]) / closes["ETH"][-4] * 100
        return "RANGING" if (b3 + e3) / 2 < -2 else "BULL"
    return "BEAR" if avg < -5 else "RANGING"


# ── Stratégies ───────────────────────────────────────────────────────────────
# Une stratégie = fonction (coin, fetch) → None ou
#   {"signal": "buy"|"sell", "mode": str, "score": float, "atr": unité de stop}
# fetch(coin, interval, n) renvoie les n dernières bougies closes ("1m","5m","15m","1h").
# Attribut STEP_MIN : fréquence d'évaluation en minutes.
def bot_strategy(bot):
    """La vraie stratégie du bot (main.get_signal), MOM et BBO autorisés quel que soit le régime."""
    def fn(coin, fetch):
        with mock.patch.object(bot, "fetch_candles", lambda c, iv, n, **k: fetch(c, iv, n)):
            s = bot.get_signal(coin)
        bot._ghost_candidates.clear()
        if not s:
            return None
        return {"signal": s["signal"], "mode": s["conditions"]["mode"], "score": s["score"],
                "atr": s["atr_m5"]}
    fn.STEP_MIN = 1
    return fn

def _strategies(bot) -> dict:
    from importlib import util
    reg = {"bot": bot_strategy(bot)}
    path = os.path.join(ROOT, "tools", "strategies.py")
    if os.path.exists(path):
        spec = util.spec_from_file_location("bst_strategies", path)
        mod = util.module_from_spec(spec); spec.loader.exec_module(mod)
        reg.update(mod.STRATEGIES)
    return reg


# ── Moteur ───────────────────────────────────────────────────────────────────
def scan(data: dict, daily: dict, strategy, bot=None) -> list[dict]:
    """Évalue la stratégie à chaque pas de temps. Retourne les signaux bruts."""
    step = getattr(strategy, "STEP_MIN", 1)
    t0 = data[COINS[0]]["1m"][0]["ts"] + 17 * 3_600_000   # chauffe : 50 bougies M15 + marge
    t0 -= t0 % (step * 60_000)
    t1 = data[COINS[0]]["1m"][-1]["ts"]
    sigs, cur = [], {}
    fetch = lambda c, iv, n: window(data, c, iv, cur["T"], n)
    patches = []
    if bot is not None:   # get_signal : pas de leçons, pas de log, MOM/BBO autorisés
        patches = [mock.patch.object(bot, "get_penalty", lambda *a, **k: 0.0),
                   mock.patch.object(bot, "log_scan", lambda *a, **k: None),
                   mock.patch.object(bot, "_current_regime", "BULL")]
    with contextlib.ExitStack() as st:
        for p in patches:
            st.enter_context(p)
        T = t0
        while T < t1:
            cur["T"] = T
            reg = None
            for c in COINS:
                s = strategy(c, fetch)
                if s:
                    reg = reg or regime_at(daily, T)
                    sigs.append({**s, "symbol": c, "T": T, "regime": reg})
            T += step * 60_000
    return sigs

def simulate(data: dict, sigs: list[dict], sl_mult=1.5, tp_mult=4.0, fee_pct=FEE_TAKER,
             max_hours=6.0, be=False, trail=False) -> list[dict]:
    """Simule les trades sur M1. Une position par coin à la fois ; entrée à l'ouverture
    de la bougie M1 du signal ; SL prioritaire si SL et TP dans la même bougie ;
    sortie au prix après max_hours. P&L en % du notionnel, frais déduits."""
    trades, busy = [], {}
    for s in sorted(sigs, key=lambda s: s["T"]):
        c = s["symbol"]
        if busy.get(c, 0) > s["T"]:
            continue
        m1, ts = data[c]["1m"], data[c]["1m_ts"]
        i = bisect.bisect_left(ts, s["T"])
        if i >= len(m1) or s["atr"] <= 0:
            continue
        entry, atr, buy = m1[i]["o"], s["atr"], s["signal"] == "buy"
        sgn = 1 if buy else -1
        sl, tp = entry - sgn * sl_mult * atr, entry + sgn * tp_mult * atr
        end_ts, protected = s["T"] + max_hours * 3_600_000, False
        out, how, j = None, "Délai", i
        for j in range(i, len(m1)):
            k = m1[j]
            if k["ts"] >= end_ts:
                out = k["o"]; break
            lo, hi = (k["l"], k["h"]) if buy else (-k["h"], -k["l"])   # côté short : prix inversés
            if lo <= sgn * sl:
                out, how = sl, ("BE" if protected else "SL"); break
            if hi >= sgn * tp:
                out, how = tp, "TP"; break
            if hi >= sgn * entry + 1.5 * atr:
                if trail:
                    sl = sgn * max(sgn * sl, hi - 1.0 * atr); protected = True
                elif be and not protected:
                    sl = sgn * max(sgn * sl, sgn * entry * (1 + sgn * 0.001)); protected = True
        if out is None:
            out = m1[-1]["c"]
        busy[c] = m1[j]["ts"] + 60_000
        trades.append({"symbol": c, "signal": s["signal"], "mode": s.get("mode", "?"),
                       "regime": s.get("regime", "?"), "score": s.get("score", 0.0), "T": s["T"],
                       "pnl": sgn * (out - entry) / entry * 100 - fee_pct, "how": how,
                       "risk_pct": sl_mult * atr / entry * 100,
                       "dur_min": (m1[j]["ts"] - s["T"]) / 60_000})
    return trades

def random_control(data: dict, sigs: list[dict], seed: int = 1, **sim) -> list[dict]:
    """Témoin : mêmes côtés et même unité de stop relative, mais coin et instant tirés au
    hasard. Si la stratégie ne le bat pas, son résultat vient du marché, pas du signal."""
    rnd = random.Random(seed)
    m1 = data[COINS[0]]["1m"]
    t0, t1 = m1[min(len(m1) - 1, 1500)]["ts"], m1[max(0, len(m1) - 500)]["ts"]
    out = []
    for s in sigs:
        c = rnd.choice(COINS)
        T = rnd.randrange(t0, t1) // 60_000 * 60_000
        px_s = window(data, s["symbol"], "1m", s["T"] + 60_000, 1)
        px_c = window(data, c, "1m", T + 60_000, 1)
        if not px_s or not px_c:
            continue
        out.append({**s, "symbol": c, "T": T, "atr": s["atr"] / px_s[0]["c"] * px_c[0]["c"]})
    return simulate(data, out, **sim)

def summary(trades: list[dict]) -> dict:
    p = [t["pnl"] for t in trades]
    if len(p) < 2:
        return {"n": len(p), "wr": 0.0, "mean": statistics.mean(p) if p else 0.0, "t": 0.0, "total": sum(p)}
    sd = statistics.stdev(p) or 1e-9
    return {"n": len(p), "wr": sum(x > 0 for x in p) / len(p) * 100, "mean": statistics.mean(p),
            "t": statistics.mean(p) / (sd / math.sqrt(len(p))), "total": sum(p)}

def _fmt(s: dict) -> str:
    return (f"n={s['n']:4}  WR={s['wr']:3.0f}%  moy={s['mean']:+.3f}%  t={s['t']:+.1f}"
            if s["n"] else "n=   0")

def report(data: dict, sigs: list[dict], **sim) -> str:
    """Rapport par mode/côté : 30 j complets, chaque moitié, et témoin aléatoire."""
    m1 = data[COINS[0]]["1m"]
    mid = (m1[0]["ts"] + m1[-1]["ts"]) // 2
    lines = [f"Paramètres : {sim}", ""]
    groups: dict[tuple, list] = {}
    for s in sigs:
        groups.setdefault((s["signal"].upper(), s.get("mode", "?")), []).append(s)
    for (side, mode), ss in sorted(groups.items()):
        tr = simulate(data, ss, **sim)
        ctrl = random_control(data, ss, **sim)
        a = summary([t for t in tr if t["T"] < mid]); b = summary([t for t in tr if t["T"] >= mid])
        regimes = {}
        for t in tr:
            regimes.setdefault(t["regime"], []).append(t)
        lines.append(f"{side} {mode}")
        lines.append(f"  période    {_fmt(summary(tr))}")
        lines.append(f"  1re moitié {_fmt(a)}")
        lines.append(f"  2e moitié  {_fmt(b)}")
        for r, rt in sorted(regimes.items()):
            lines.append(f"  {r:10} {_fmt(summary(rt))}")
        lines.append(f"  témoin     {_fmt(summary(ctrl))}  (entrées au hasard)")
        lines.append("")
    return "\n".join(lines)


def main_cli(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--days", type=int, default=30)
    ap.add_argument("--strategy", default="bot")
    ap.add_argument("--sl", type=float, default=1.5, help="SL en multiples de l'unité de stop (ATR)")
    ap.add_argument("--tp", type=float, default=4.0, help="TP en multiples de l'unité de stop (ATR)")
    ap.add_argument("--fee", type=float, default=FEE_TAKER, help="frais aller-retour en %%")
    ap.add_argument("--max-hours", type=float, default=6.0)
    ap.add_argument("--be", action="store_true", help="break-even à 1.5×ATR")
    ap.add_argument("--trail", action="store_true", help="trailing à 1×ATR dès 1.5×ATR")
    ap.add_argument("--refresh", action="store_true", help="re-télécharger les bougies")
    ap.add_argument("--no-download", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args(argv)

    bot = load_bot()
    strategies = _strategies(bot)
    if a.list:
        for k, f in strategies.items():
            print(f"{k:14} {(f.__doc__ or '').strip().splitlines()[0] if f.__doc__ else ''}")
        return
    if a.strategy not in strategies:
        sys.exit(f"Stratégie inconnue : {a.strategy} (voir --list)")
    if not a.no_download:
        download(a.days, a.refresh)
    data, daily = load(a.days)

    cache = os.path.join(CACHE_DIR, f"sigs_{a.strategy}_{a.days}d_{data[COINS[0]]['1m'][-1]['ts']}"
                                    f"_vm{os.environ.get('VOL_MULTIPLIER')}.pkl")
    if os.path.exists(cache):
        with open(cache, "rb") as f:
            sigs = pickle.load(f)
    else:
        t = time.time()
        sigs = scan(data, daily, strategies[a.strategy], bot if a.strategy == "bot" else None)
        with open(cache, "wb") as f:
            pickle.dump(sigs, f)
        print(f"  scan : {len(sigs)} signaux en {time.time() - t:.0f}s", file=sys.stderr)
    print(report(data, sigs, sl_mult=a.sl, tp_mult=a.tp, fee_pct=a.fee,
                 max_hours=a.max_hours, be=a.be, trail=a.trail))


if __name__ == "__main__":
    main_cli()
