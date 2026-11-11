#!/usr/bin/env python3
"""
Backtest — stratégie tendance journalière de BotShortTrade (main.trend_cycle)
============================================================================
Rejoue la règle du bot, avec ses paramètres et ses coins (lus dans main.py), sur l'historique journalier :
  - LONG si la clôture D1 dépasse celle d'il y a N jours pour la majorité des horizons
    TREND_LOOKBACKS (14, 28, 56), SHORT sinon ;
  - quantité fixée à l'ouverture (TREND_EXPOSURE × équité / nb de coins), tenue jusqu'à l'inversion ;
  - SL à TREND_SL_ATR × ATR14 D1, puis à plat jusqu'à la clôture journalière suivante ;
  - frais taker 0.0432 % par côté ; funding Bybit (avant son historique : longs 15 %/an, shorts 0).
Données publiques mises en cache dans backtest_data/ : bougies D1 Binance (depuis 2017) + funding Bybit.

Usage :
    python3 tools/trend_backtest.py                # rapport par période
    python3 tools/trend_backtest.py --perm 500     # + témoin par permutation (lent)
    python3 tools/trend_backtest.py --expo 1.0     # autre exposition

Témoin « permutation » : le signal de chaque coin est décalé d'un nombre aléatoire de jours (même
proportion long/short, même rythme d'inversions, mais plus aligné sur les prix). p = part des tirages
au moins aussi bons que la vraie stratégie ; un signal sans valeur donne p ≈ 0.5.

Résultats du 2026-11-02 (5 coins, exposition 50 %) :
  - 28 j seul (règle initiale) : 2017→2026 +35 %/an, Sharpe 0.86, pire baisse −49 %, permutation p=0.02 ;
    2021→2026 +12 %/an, pire baisse −32 % ; dernière année ≈ 0 %.
  - vote 14/28/56 j (adopté) : 2017→2026 +40 %/an, Sharpe 0.94, pire baisse −49 %, p=0.01 ;
    1re moitié Sharpe 1.18 (vs 1.10), 2e moitié 0.63 (vs 0.54), pire baisse 2e moitié −41 % (vs −32 %).
    Écart avec la règle 28 j : +3.8 %/an, t=0.9 (non significatif) ; adopté pour la robustesse.
  - rejetés (même protocole, hypothèses déclarées avant les tests) : long/plat (moins bon partout),
    10 coins (2e moitié moins bonne), taille ∝ 1/volatilité (1re moitié moins bonne).
Résultats du 2026-11-03 (2e série, mêmes critères) :
  - BTC à la place de LTC (adopté ; BTC = 1er perp Hyperliquid en volume, LTC = 40e) : 2017→2026
    +50 %/an, Sharpe 1.11, pire baisse −45 %, p=0.003 ; 1re moitié 1.27 (vs 1.18), 2e moitié 0.99 (vs 0.63),
    pire baisse 2e moitié −34 % (vs −41 %). Écart avec LTC : +6.4 %/an, t=2.2.
  - rejetés : décision hebdomadaire (moins bon partout), filtre de funding extrême (moins bon partout).
Résultats du 2026-11-04 (3e série, techniques documentées chez les traders, mêmes critères) :
  - exposition selon la volatilité (Harvey et al. 2018) : taille × médiane passée de σ60 / σ60 du jour,
    bornée à [0.5, 2], même facteur pour les 5 coins. Sharpe 1.16 (vs 1.11), 1re moitié 1.19 (vs 1.27),
    2e moitié 1.19 (vs 0.99), pire baisse −38 % (vs −45 %). Rejeté : moins bon sur la 1re moitié
    (exposition moyenne 57 %, pas 50 %) ; à revoir sur des données nouvelles.
  - stop suiveur « chandelier » (plus haut depuis l'entrée − 3×ATR14) : +40 %/an (vs +50 %), Sharpe 1.21,
    1re moitié 1.39, 2e moitié 0.98 (vs 0.99), pire baisse −32 % (vs −45 %). Rejeté : moins de gain, et
    pas mieux sur la 2e moitié.
  - écartés sans test : carry de funding (aucun marché spot liquide sur Hyperliquid pour couvrir), taille
    continue selon la force du signal (Carver) : impossible tant que chaque coin est au minimum de 10,5 $.
Résultats du 2026-11-04 (4e série, techniques plus complexes, mêmes critères) :
  - ensemble EWMAC de Carver (8/32 à 64/256, signe de la moyenne des prévisions plafonnées) : dès 2018-05,
    Sharpe 0.80 (vs 1.09), 2e moitié 0.31 (vs 0.84), pire baisse −72 % (vs −45 %). Rejeté : trop lent pour
    le crypto (194 inversions contre 1 236).
  - méta-étiquetage (López de Prado) : régression logistique réentraînée chaque année sur le passé seul,
    entrée si P(gagnante à 14 j) ≥ 0.5. Justesse hors échantillon 49 % pour 51 % de positions gagnantes :
    aucun pouvoir de prédiction ; 2020→2026 Sharpe 0.97 (vs 1.18). Rejeté.
  - la règle du bot passe le ratio de Sharpe « dégonflé » (Bailey et López de Prado) même pour 30 essais,
    mais son résultat tient à quelques journées extrêmes (DOGE le 2021-01-28 : +75 % d'équité en un
    jour) : Sharpe 1.11 → 0.78 sans les 10 meilleures journées sur 3 300, 2e moitié 1.00 → 0.49. Normal
    pour un suivi de tendance ; ne jamais filtrer les entrées ni laisser le bot à l'arrêt.
Résultats du 2026-11-04 (5e série, diversification multi-actifs comme les fonds de tendance, mêmes critères) :
  - ajout de 3 marchés Hyperliquid (trade.xyz, marge croisée) : S&P 500, or, pétrole Brent (historique Yahoo
    depuis 2015, week-ends à plat), chacun au même risque qu'un coin crypto (ATR% médian crypto / ATR% du
    marché, plafond 5× ; notionnel moyen des 3 : 121 % de l'équité). Coûts prudents : frais 0.045 %/côté
    (or 0.09 %), longs paient 5.5 %/an de funding, shorts ne touchent rien. Sharpe 0.93 (vs 1.11),
    1re moitié 1.16 (vs 1.27), 2e moitié 0.65 (vs 0.97), pire baisse −46 % (vs −45 %). Rejeté.
  - avec les frais actuels (growth mode 0.009 %, or exclu) et le funding reçu par les shorts : 1.01 / 1.21 /
    0.77, toujours moins bon. Les 3 marchés seuls : −5.7 %/an (−2.1 %/an avec ces coûts), corrélation 0.08
    avec les cryptos. La règle 14/28/56 j est trop rapide pour eux : l'or, triplé depuis 2017, est short 42 %
    du temps et ne gagne rien avant frais ; le funding coûte et la marge en USDC ne rapporte rien.
"""
from __future__ import annotations

import argparse, contextlib, io, json, math, os, random, statistics, sys, tempfile, time, types, urllib.request

ROOT       = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR  = os.path.join(ROOT, "backtest_data")

FEE        = 0.000432
BINANCE    = "https://data-api.binance.vision/api/v3/klines"
BYBIT_FUND = "https://api.bybit.com/v5/market/funding/history"
DAY_MS     = 86_400_000


def load_bot():
    """Importe main.py (paramètres et coins de la stratégie) sans .env, sans SDK réel, dans un dossier temporaire :
    aucun ordre possible."""
    if "main" in sys.modules:          # déjà importé (ex. par tests/) → réutiliser
        return sys.modules["main"]
    dotenv = types.ModuleType("dotenv"); dotenv.load_dotenv = lambda *a, **k: None
    sys.modules["dotenv"] = dotenv
    os.environ.update({"HL_PRIVATE_KEY": "0x" + "11" * 32, "HL_TESTNET": "1", "DISCORD_WEBHOOK_URL": ""})
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

def _get(url: str):
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=20) as r:
                return json.load(r)
        except Exception:
            time.sleep(1 + attempt)
    raise RuntimeError(f"injoignable : {url}")

def download(coins: list[str], refresh: bool = False) -> None:
    """Bougies D1 Binance et funding Bybit de chaque coin (cache de moins d'un jour réutilisé)."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    for coin in coins:
        path = os.path.join(CACHE_DIR, f"{coin}_trend_1d.json")
        if not refresh and os.path.exists(path) and time.time() - os.path.getmtime(path) < DAY_MS / 1000:
            continue
        rows, start = {}, 0
        while True:
            lst = _get(f"{BINANCE}?symbol={coin}USDT&interval=1d&limit=1000&startTime={start}")
            for k in lst:
                rows[int(k[0])] = [float(x) for x in k[1:5]]
            if len(lst) < 1000:
                break
            start = int(lst[-1][0]) + 1
            time.sleep(0.2)
        fund, end = {}, int(time.time() * 1000)
        while True:
            lst = _get(f"{BYBIT_FUND}?category=linear&symbol={coin}USDT&limit=200&endTime={end}")["result"]["list"]
            for r in lst:
                fund[int(r["fundingRateTimestamp"])] = float(r["fundingRate"])
            if len(lst) < 200:
                break
            end = min(int(r["fundingRateTimestamp"]) for r in lst) - 1
            time.sleep(0.15)
        with open(path, "w") as f:
            json.dump({"d1": [[k] + rows[k] for k in sorted(rows)], "funding": sorted(fund.items())}, f)
        print(f"  {coin} : {len(rows)} jours, {len(fund)} funding", file=sys.stderr)

def load(coins: list[str]) -> dict:
    """{coin: {"days": [...], "ohlc": {jour: (o, h, l, c)}, "fund": {jour: taux cumulé}, "fund_start": jour}}"""
    data = {}
    for coin in coins:
        with open(os.path.join(CACHE_DIR, f"{coin}_trend_1d.json")) as f:
            raw = json.load(f)
        fund: dict[int, float] = {}
        for ts, r in raw["funding"]:
            fund[ts // DAY_MS * DAY_MS] = fund.get(ts // DAY_MS * DAY_MS, 0.0) + r
        ohlc = {r[0]: tuple(r[1:5]) for r in raw["d1"]}
        data[coin] = {"days": sorted(ohlc), "ohlc": ohlc, "fund": fund,
                      "fund_start": min(fund) if fund else float("inf")}
    return data

def signals(data: dict, lookback: int) -> dict[str, dict[int, int]]:
    """Signal décidé à la clôture de chaque jour : +1 long, −1 short."""
    out = {}
    for coin, d in data.items():
        cl = [d["ohlc"][t][3] for t in d["days"]]
        out[coin] = {d["days"][i]: (1 if cl[i] > cl[i - lookback] else -1) for i in range(lookback, len(cl))}
    return out

def vote(data: dict, lookbacks) -> dict[str, dict[int, int]]:
    """Vote majoritaire des signaux de chaque horizon (règle du bot : égalité → short)."""
    ss = [signals(data, lb) for lb in lookbacks]
    return {c: {t: (1 if sum(s[c][t] for s in ss) > 0 else -1) for t in ss[-1][c] if all(t in s[c] for s in ss)}
            for c in data}

def shifted(sig: dict[str, dict[int, int]], rnd: random.Random) -> dict[str, dict[int, int]]:
    """Témoin : chaque série de signaux décalée circulairement d'un nombre aléatoire de jours."""
    out = {}
    for coin, s in sig.items():
        days = sorted(s)
        vals = [s[t] for t in days]
        k = rnd.randrange(1, len(vals))
        out[coin] = dict(zip(days, vals[k:] + vals[:k]))
    return out

def simulate(data: dict, sig: dict, sl_atr: float, expo: float, start: int | None = None,
             end: int | None = None) -> list[tuple[int, float]]:
    """Courbe d'équité (jour, équité) de la règle du bot, équité initiale 1."""
    coins = list(data)
    days  = sorted(set().union(*[set(d["days"]) for d in data.values()]))
    E, pos, curve = 1.0, {}, []
    hist: dict[str, list[tuple]] = {c: [] for c in coins}
    for day in days:
        live = (start is None or day >= start) and (end is None or day < end)
        for c in coins:                       # journée : SL, valorisation, funding
            bar = data[c]["ohlc"].get(day)
            if not bar:
                continue
            hist[c].append(bar)
            p = pos.get(c)
            if not p or not live:
                continue
            o, h, l, cl = bar
            if (p["s"] > 0 and l <= p["sl"]) or (p["s"] < 0 and h >= p["sl"]):
                gap = (p["s"] > 0 and o < p["sl"]) or (p["s"] < 0 and o > p["sl"])
                px = o if gap else p["sl"]
                E += p["s"] * p["q"] * (px - p["last"]) - FEE * p["q"] * px
                del pos[c]
                continue
            f = data[c]["fund"].get(day, 0.15 / 365 if day < data[c]["fund_start"] else 0.0)
            if day < data[c]["fund_start"] and p["s"] < 0:
                f = 0.0                           # prudent : les shorts ne touchent rien avant Bybit
            E += p["s"] * p["q"] * (cl - p["last"]) - p["s"] * f * p["q"] * cl
            p["last"] = cl
        if not live:
            continue
        for c in coins:                       # clôture : inversion → fermeture, puis ouverture
            want, hh = sig[c].get(day), hist[c]
            if want is None or day not in data[c]["ohlc"] or len(hh) < 15:
                continue
            cl, p = hh[-1][3], pos.get(c)
            if p and p["s"] != want:
                E -= FEE * p["q"] * cl
                del pos[c]
                p = None
            if not p:
                atr = sum(max(hh[j][1] - hh[j][2], abs(hh[j][1] - hh[j - 1][3]), abs(hh[j][2] - hh[j - 1][3]))
                          for j in range(-14, 0)) / 14
                q = expo * E / len(coins) / cl
                E -= FEE * q * cl
                pos[c] = {"s": want, "q": q, "last": cl, "sl": cl - want * sl_atr * atr}
        curve.append((day, E))
    return curve

def stats(curve: list[tuple[int, float]]) -> dict:
    rets = [b[1] / a[1] - 1 for a, b in zip(curve, curve[1:])]
    mu, sd = statistics.mean(rets), statistics.stdev(rets)
    peak, dd = 0.0, 0.0
    for _, e in curve:
        peak = max(peak, e)
        dd = min(dd, e / peak - 1)
    years = len(rets) / 365
    return {"cagr": (curve[-1][1] / curve[0][1]) ** (1 / years) - 1, "sharpe": mu / sd * math.sqrt(365),
            "t": mu / (sd / math.sqrt(len(rets))), "dd": dd, "mean": mu}

def _fmt(label: str, curve: list) -> str:
    s = stats(curve)
    span = "→".join(time.strftime("%Y-%m", time.gmtime(t / 1000)) for t in (curve[0][0], curve[-1][0]))
    return (f"{label:24s} {span}  {s['cagr'] * 100:+6.1f} %/an  sharpe={s['sharpe']:+.2f}  "
            f"t={s['t']:+.2f}  pire baisse={s['dd'] * 100:5.1f} %")


def main_cli(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--expo", type=float, default=None, help="exposition (défaut : TREND_EXPOSURE du bot)")
    ap.add_argument("--perm", type=int, default=0, help="nb de tirages du témoin permutation")
    ap.add_argument("--refresh", action="store_true")
    ap.add_argument("--no-download", action="store_true")
    a = ap.parse_args(argv)

    bot = load_bot()
    coins, expo = list(bot.TREND_COINS), a.expo if a.expo is not None else bot.TREND_EXPOSURE
    if not a.no_download:
        download(coins, a.refresh)
    data = load(coins)
    sig  = vote(data, bot.TREND_LOOKBACKS)
    sim  = lambda s, **k: simulate(data, s, bot.TREND_SL_ATR, expo, **k)
    now  = int(time.time() * 1000)
    print(f"Tendance (vote {'/'.join(map(str, bot.TREND_LOOKBACKS))} j) · SL {bot.TREND_SL_ATR:g}×ATR · "
          f"exposition {expo:.0%} · {', '.join(coins)}\n")
    full = sim(sig)
    print(_fmt("toute la période", full))
    for label, start in (("5 dernières années", now - 5 * 365 * DAY_MS),
                         ("2 dernières années", now - 2 * 365 * DAY_MS),
                         ("dernière année", now - 365 * DAY_MS)):
        print(_fmt(label, sim(sig, start=start)))
    mid = full[len(full) // 2][0]
    print(_fmt("1re moitié", sim(sig, end=mid)))
    print(_fmt("2e moitié", sim(sig, start=mid)))
    inv = {c: {t: -v for t, v in s.items()} for c, s in sig.items()}
    print(_fmt("témoin inversé", sim(inv)))
    if a.perm:
        rnd, obs, ge = random.Random(7), stats(full)["mean"], 0
        for _ in range(a.perm):
            ge += stats(sim(shifted(sig, rnd)))["mean"] >= obs
        print(f"témoin permutation       p = {(ge + 1) / (a.perm + 1):.3f} ({a.perm} tirages)")


if __name__ == "__main__":
    main_cli()
