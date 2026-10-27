"""
Stratégies alternatives pour tools/backtest.py (recherche — aucune n'est utilisée par le bot).

Chaque stratégie : fn(coin, fetch) → None ou {"signal", "mode", "score", "atr"}
  fetch(coin, interval, n) = n dernières bougies closes ("1m", "5m", "15m", "1h")
  "atr" = unité de stop : SL = --sl × atr, TP = --tp × atr
  STEP_MIN = fréquence d'évaluation (minutes)

Idée commune : viser des mouvements bien plus grands que les frais (~0.09 % aller-retour),
donc des unités de temps plus longues que le M5 du bot.
"""
from __future__ import annotations


def _ema(xs: list[float], n: int) -> list[float]:
    k, out = 2 / (n + 1), [xs[0]]
    for x in xs[1:]:
        out.append(x * k + out[-1] * (1 - k))
    return out

def _rsi(xs: list[float], n: int = 14) -> float:
    gains = [max(0.0, b - a) for a, b in zip(xs, xs[1:])]
    losses = [max(0.0, a - b) for a, b in zip(xs, xs[1:])]
    ag, al = sum(gains[:n]) / n, sum(losses[:n]) / n
    for g, l in zip(gains[n:], losses[n:]):           # lissage de Wilder
        ag, al = (ag * (n - 1) + g) / n, (al * (n - 1) + l) / n
    return 100.0 if al == 0 else 100 - 100 / (1 + ag / al)

def _atr(cs: list[dict], n: int = 14) -> float:
    trs = [max(c["h"] - c["l"], abs(c["h"] - p["c"]), abs(c["l"] - p["c"])) for p, c in zip(cs, cs[1:])]
    return sum(trs[-n:]) / min(n, len(trs)) if trs else 0.0

def _bb(xs: list[float], n: int = 20, k: float = 2.0) -> tuple[float, float, float]:
    w = xs[-n:]
    mu = sum(w) / n
    sd = (sum((x - mu) ** 2 for x in w) / n) ** 0.5
    return mu - k * sd, mu, mu + k * sd


def h1_trend(coin, fetch):
    """Suivi de tendance H1 : cassure du plus haut/bas 20 h dans le sens de l'EMA50 H1."""
    cs = fetch(coin, "1h", 80)
    if len(cs) < 60:
        return None
    closes = [c["c"] for c in cs]
    ema50 = _ema(closes, 50)[-1]
    hi20 = max(c["h"] for c in cs[-21:-1]); lo20 = min(c["l"] for c in cs[-21:-1])
    last = cs[-1]["c"]
    atr = _atr(cs)
    if last > hi20 and last > ema50:
        return {"signal": "buy", "mode": "H1T", "score": 1.0, "atr": atr}
    if last < lo20 and last < ema50:
        return {"signal": "sell", "mode": "H1T", "score": 1.0, "atr": atr}
    return None
h1_trend.STEP_MIN = 60


def h1_meanrev(coin, fetch):
    """Retour à la moyenne H1 : clôture hors bande de Bollinger 20/2 avec RSI14 extrême."""
    cs = fetch(coin, "1h", 60)
    if len(cs) < 40:
        return None
    closes = [c["c"] for c in cs]
    lo, _, hi = _bb(closes)
    rsi = _rsi(closes)
    atr = _atr(cs)
    if closes[-1] < lo and rsi < 30:
        return {"signal": "buy", "mode": "H1MR", "score": 1.0, "atr": atr}
    if closes[-1] > hi and rsi > 70:
        return {"signal": "sell", "mode": "H1MR", "score": 1.0, "atr": atr}
    return None
h1_meanrev.STEP_MIN = 60


def m15_pullback(coin, fetch):
    """Tendance M15 (EMA50 > EMA200, pente positive) + repli sous l'EMA20 puis reprise :
    on rejoint la tendance après une correction (et l'inverse pour la vente)."""
    cs = fetch(coin, "15m", 230)
    if len(cs) < 220:
        return None
    closes = [c["c"] for c in cs]
    e20, e50, e200 = _ema(closes, 20), _ema(closes, 50), _ema(closes, 200)
    atr = _atr(cs)
    up = e50[-1] > e200[-1] and e50[-1] > e50[-10]
    dn = e50[-1] < e200[-1] and e50[-1] < e50[-10]
    # repli : la bougie précédente a clôturé de l'autre côté de l'EMA20, la dernière revient
    if up and closes[-2] < e20[-2] and closes[-1] > e20[-1] and _rsi(closes) < 60:
        return {"signal": "buy", "mode": "M15PB", "score": 1.0, "atr": atr}
    if dn and closes[-2] > e20[-2] and closes[-1] < e20[-1] and _rsi(closes) > 40:
        return {"signal": "sell", "mode": "M15PB", "score": 1.0, "atr": atr}
    return None
m15_pullback.STEP_MIN = 15


STRATEGIES = {"h1_trend": h1_trend, "h1_meanrev": h1_meanrev, "m15_pullback": m15_pullback}
