"""
Stratégie tendance sur un Hyperliquid simulé avec ses vraies règles (tailles au pas du coin, prix à
5 chiffres significatifs, SL reduce-only), pendant des mois de prix aléatoires sur les 5 coins :
  - chaque ordre (ouverture, SL, renfort, fermeture) vaut au moins TREND_MIN_NOTIONAL à son pire prix ;
  - après chaque cycle, les positions suivies par le bot sont exactement celles du compte, chacune
    avec son SL ; aucun SL orphelin ;
  - le P&L et le prix de sortie enregistrés de chaque position fermée sont ceux du compte (fills).
Aucun réseau : test_regression (importé en premier) isole main.py.
"""
import math, random, time, unittest
from unittest import mock

import test_regression as T          # isolation (stubs, réseau interdit, dossier temporaire) EN PREMIER
m = T.m

SZ_DEC = {"BTC": 5, "ETH": 4, "SOL": 2, "XRP": 0, "DOGE": 0}               # Hyperliquid, 2026-11-04
START  = {"BTC": 78491.0, "ETH": 2324.0, "SOL": 93.193, "XRP": 1.2841, "DOGE": 0.075238}


def _hl_px(px: float, coin: str) -> float:
    """Prix au format Hyperliquid, arrondi comme le SDK (5 chiffres significatifs, 6 − szDecimals décimales)."""
    return round(float(f"{px:.5g}"), 6 - SZ_DEC[coin])


class HLSim:
    """Compte Hyperliquid simulé : ordres au marché exécutés au prix moyen, SL déclenchés quand le prix
    les franchit, fills avec closedPnl et frais comme l'API. Toute règle enfreinte va dans `problems`."""
    def __init__(self, fail_close: float = 0.0, rnd: random.Random | None = None):
        self.mids     = dict(START)
        self.pos: dict[str, dict] = {}       # coin → {"szi", "entry"}
        self.stops: dict[int, dict] = {}     # oid → SL au repos
        self.fills: list[dict] = []
        self.problems: list[str] = []
        self.done: dict[str, list[tuple]] = {}   # coin → (P&L réel, prix de sortie) de chaque position fermée
        self._ep: dict[str, float] = {}
        self.fail_close, self.rnd = fail_close, rnd or random.Random(0)
        self.oid, self.t = 100, time.time() * 1000

    def _next_oid(self) -> int:
        self.oid += 1
        return self.oid

    def _check(self, kind: str, coin: str, sz: float, limit: float, trigger: float | None = None):
        dec = SZ_DEC[coin]
        if sz <= 0 or abs(round(sz, dec) - sz) > 1e-12:
            self.problems.append(f"{kind} {coin} : taille {sz} hors du pas 1e-{dec}")
        for px in (limit, trigger):
            if px is not None and not T._hl_px_valide(px, dec):
                self.problems.append(f"{kind} {coin} : prix {px} au mauvais format")
        worst = min(limit, self.mids[coin] if trigger is None else trigger)
        if sz * worst < m.TREND_MIN_NOTIONAL - 1e-9:
            self.problems.append(f"{kind} {coin} : {sz} × {worst:.6g} = {sz * worst:.2f} $ "
                                 f"< {m.TREND_MIN_NOTIONAL:g} $")

    def _fill(self, coin: str, is_buy: bool, sz: float, px: float, oid: int):
        p   = self.pos.setdefault(coin, {"szi": 0.0, "entry": 0.0})
        sgn = 1 if is_buy else -1
        if p["szi"] == 0 or (p["szi"] > 0) == is_buy:      # ouverture ou renfort
            new, pnl = p["szi"] + sgn * sz, 0.0
            p["entry"] = (abs(p["szi"]) * p["entry"] + sz * px) / abs(new)
            d = "Open Long" if is_buy else "Open Short"
        else:                                               # réduction (reduce-only)
            sz  = min(sz, abs(p["szi"]))
            pnl = (px - p["entry"]) * sz * (1 if p["szi"] > 0 else -1)
            new = p["szi"] + sgn * sz
            d   = "Close Long" if p["szi"] > 0 else "Close Short"
        p["szi"] = round(new, SZ_DEC[coin])
        fee = m.TREND_FEE * sz * px
        self.t = max(self.t + 1, time.time() * 1000)
        self.fills.append({"coin": coin, "dir": d, "time": self.t, "oid": oid, "px": str(px), "sz": str(sz),
                           "closedPnl": str(pnl), "fee": str(fee), "side": "B" if is_buy else "A"})
        self._ep[coin] = self._ep.get(coin, 0.0) + pnl - fee
        if p["szi"] == 0:
            del self.pos[coin]
            self.done.setdefault(coin, []).append((self._ep.pop(coin), px))

    @staticmethod
    def _filled(sz: float, px: float, oid: int) -> dict:
        return {"status": "ok", "response": {"type": "order", "data": {"statuses": [
            {"filled": {"totalSz": str(sz), "avgPx": str(px), "oid": oid}}]}}}

    # ── Exchange ──
    def update_leverage(self, *a, **k):
        return {"status": "ok"}

    def market_open(self, coin, is_buy, sz, px=None, slippage=0.05):
        mid = self.mids[coin]
        self._check("ouverture", coin, sz, _hl_px(mid * (1 + slippage if is_buy else 1 - slippage), coin))
        oid = self._next_oid()
        self._fill(coin, is_buy, sz, mid, oid)
        return self._filled(sz, mid, oid)

    def market_close(self, coin, sz=None, px=None, slippage=0.05):
        p = self.pos.get(coin)
        if not p:
            return None
        is_buy, sz, mid = p["szi"] < 0, abs(p["szi"]), self.mids[coin]
        self._check("fermeture", coin, sz, _hl_px(mid * (1 + slippage if is_buy else 1 - slippage), coin))
        if self.rnd.random() < self.fail_close:
            return {"status": "ok", "response": {"type": "order", "data": {"statuses": [
                {"error": "Order could not immediately match against any resting orders."}]}}}
        oid = self._next_oid()
        self._fill(coin, is_buy, sz, mid, oid)
        return self._filled(sz, mid, oid)

    def order(self, coin, is_buy, sz, limit_px, order_type, reduce_only):
        trig = order_type["trigger"]
        self._check("SL", coin, sz, limit_px, trig["triggerPx"])
        p = self.pos.get(coin)
        if not reduce_only or not p or (p["szi"] < 0) != is_buy or sz > abs(p["szi"]) + 1e-12:
            self.problems.append(f"SL {coin} : pas un reduce-only opposé à la position ({sz}, {p})")
        oid = self._next_oid()
        self.stops[oid] = {"coin": coin, "is_buy": is_buy, "sz": sz, "limit": limit_px, "trigger": trig["triggerPx"]}
        return {"status": "ok", "response": {"type": "order", "data": {"statuses": [{"resting": {"oid": oid}}]}}}

    def cancel(self, coin, oid):
        if self.stops.pop(oid, None) is None:
            return {"status": "err", "response": "Order was never placed, already canceled, or filled."}
        return {"status": "ok", "response": {"type": "cancel", "data": {"statuses": ["success"]}}}

    # ── Info ──
    def all_mids(self):
        return {c: str(px) for c, px in self.mids.items()}

    def user_state(self, addr=None):
        return {"assetPositions": [{"position": {"coin": c, "szi": str(p["szi"]), "entryPx": str(p["entry"]),
                                                 "unrealizedPnl": str((self.mids[c] - p["entry"]) * p["szi"])}}
                                   for c, p in self.pos.items()]}

    def user_fills(self, addr=None):
        return list(reversed(self.fills))            # plus récent d'abord, comme l'API

    def frontend_open_orders(self, addr=None):
        return [{"coin": o["coin"], "oid": oid, "isTrigger": True, "reduceOnly": True, "orderType": "Stop Market",
                 "triggerPx": str(o["trigger"]), "sz": str(o["sz"])} for oid, o in self.stops.items()]

    # ── Marché ──
    def move(self, coin: str, px: float):
        """Nouveau prix ; déclenche les SL franchis (exécutés au prix, borné par leur limite)."""
        self.mids[coin] = px
        for oid, o in list(self.stops.items()):
            if o["coin"] == coin and (px >= o["trigger"] if o["is_buy"] else px <= o["trigger"]):
                del self.stops[oid]
                p = self.pos.get(coin)
                if p and (p["szi"] < 0) == o["is_buy"]:
                    fill = min(px, o["limit"]) if o["is_buy"] else max(px, o["limit"])
                    self._fill(coin, o["is_buy"], min(o["sz"], abs(p["szi"])), fill, oid)


class TestTrendSurHyperliquidSimule(T.BaseTest):
    def simulate(self, seed: int, equity: float, days: int = 150, ticks: int = 4, vol: float = 0.05,
                 fail_close: float = 0.0, paper_days: int = 0, drift: float = 0.0):
        """`days` jours de prix aléatoires (`ticks` cycles par jour), dont les `paper_days` premiers en
        mode fantôme ; retourne le compte simulé et les positions fermées."""
        rnd = random.Random(seed)
        sim = HLSim(fail_close, rnd)
        hist = {}
        for c in m.TREND_COINS:                              # 60 jours d'historique finissant au prix de départ
            walk = [0.0]
            for _ in range(59):
                walk.append(walk[-1] + rnd.gauss(drift, vol))
            hist[c] = [START[c] * math.exp(w - walk[-1]) for w in walk]
        wb, day, closed = m.init_excel(), [0], []
        with mock.patch.object(m, "info", sim), mock.patch.object(m, "exchange", sim), \
             mock.patch.object(m, "fetch_candles", lambda coin, iv, n, **k: T._d1(hist[coin][-n:])), \
             mock.patch.object(m, "GHOST_ONLY", False), mock.patch.object(m, "SL_CHECK_S", 0), \
             mock.patch.object(m, "_utc_day", lambda ts: f"J{day[0]:04d}"), \
             mock.patch.object(m, "_safe_save", lambda wb: None), \
             mock.patch.dict(m._sym_filters, {c: {"sz_dec": d, "sz_step": 10 ** -d} for c, d in SZ_DEC.items()}):
            for day[0] in range(days):
                m.GHOST_ONLY = day[0] < paper_days          # remis à sa valeur par le patch en sortie
                for _ in range(ticks):
                    for c in m.TREND_COINS:
                        sim.move(c, _hl_px(sim.mids[c] * math.exp(rnd.gauss(drift / ticks, vol / ticks ** 0.5)), c))
                    closed += m.trend_cycle(wb, sim.user_state(), equity)
                    self.assert_compte_coherent(sim)
                for c in m.TREND_COINS:
                    hist[c].append(sim.mids[c])
        self.assertEqual(sim.problems, [])
        bot = {}
        for p in closed:
            if not p.get("paper"):
                bot.setdefault(p["symbol"], []).append((p["pnl"], p["exit"]))
        for c in m.TREND_COINS:              # P&L et sortie enregistrés = ceux du compte, position par position
            self.assertEqual(len(bot.get(c, [])), len(sim.done.get(c, [])), c)
            for (pnl, exit_px), (real_pnl, real_px) in zip(bot.get(c, []), sim.done.get(c, [])):
                self.assertAlmostEqual(pnl, real_pnl, places=3, msg=c)
                self.assertAlmostEqual(exit_px, real_px, places=9, msg=c)
        return sim, closed

    def assert_compte_coherent(self, sim: HLSim):
        poss = {c: p for c, p in m._load_trend()["positions"].items() if not p.get("paper")}
        self.assertEqual(set(poss), set(sim.pos), "positions du bot ≠ positions du compte")
        for c, p in poss.items():
            hl = sim.pos[c]
            self.assertEqual(hl["szi"] > 0, p["side"] == "BUY", c)
            self.assertAlmostEqual(abs(hl["szi"]), p["qty"] + p.get("top_up", 0.0), places=9, msg=c)
            stops = {oid: o for oid, o in sim.stops.items() if o["coin"] == c}
            self.assertEqual(list(stops), [p["sl_oid"]], f"{c} : un seul SL, celui que suit le bot")
            # Le SL couvre la quantité ouverte ; un renfort n'existe que le temps de réessayer sa fermeture
            self.assertAlmostEqual(stops[p["sl_oid"]]["sz"], p["qty"], places=9, msg=c)
        self.assertEqual({o["coin"] for o in sim.stops.values()} - set(poss), set(), "SL orphelin")

    def statuses(self, closed: list[dict]) -> set[str]:
        return {p["status"] for p in closed}

    def test_capital_actuel_111_dollars(self):
        """Une semaine en fantôme puis 5 mois en réel : passage en réel, inversions, SL touchés et
        renforts de shorts très gagnants (fermeture seule < 10,5 $), tous les ordres valides."""
        sim, closed = self.simulate(1, 111.0, paper_days=7)
        self.assertLessEqual({"fin du mode fantôme", "signal inversé", "SL touché"}, self.statuses(closed))
        self.assertTrue(any(p.get("top_up") for p in closed))

    def test_capital_conseille_200_dollars(self):
        sim, closed = self.simulate(2, 200.0)
        self.assertLessEqual({"signal inversé", "SL touché"}, self.statuses(closed))

    def test_gros_capital_1000_dollars(self):
        sim, closed = self.simulate(3, 1000.0)
        self.assertLessEqual({"signal inversé", "SL touché"}, self.statuses(closed))
        self.assertFalse(any(p.get("top_up") for p in closed))     # positions de 100 $ : jamais de renfort

    def test_fermetures_refusees_au_hasard(self):
        """Une fermeture sur trois refusée : nouvel essai au cycle suivant, jamais de position oubliée
        ni de renforts en série, P&L toujours celui du compte."""
        sim, closed = self.simulate(5, 111.0, fail_close=0.33)
        self.assertGreater(len(closed), 30)


if __name__ == "__main__":
    unittest.main()
