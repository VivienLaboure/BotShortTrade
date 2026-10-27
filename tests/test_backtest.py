"""
Tests de l'outil de backtest (tools/backtest.py) — sans réseau.
"""
import os, sys, unittest
from unittest import mock

import test_regression as T          # isolation de main.py (stubs, dossier temporaire) EN PREMIER
m = T.m
sys.path.insert(0, os.path.join(T.ROOT, "tools"))
import backtest as B                  # noqa: E402
import strategies as S                # noqa: E402

H = 3_600_000


def _m1(n, start=0, px=100.0, step=0.0, vol=10.0):
    return [{"ts": start + i * 60_000, "o": px + step * i, "h": px + step * i + 0.05,
             "l": px + step * i - 0.05, "c": px + step * (i + 1), "v": vol} for i in range(n)]

def _data(m1_by_coin):
    data = {}
    for c, m1 in m1_by_coin.items():
        d = {"1m": m1, "5m": B.agg(m1, 5), "15m": B.agg(m1, 15), "1h": B.agg(m1, 60)}
        for iv in list(d):
            d[iv + "_ts"] = [x["ts"] for x in d[iv]]
        data[c] = d
    return data

def _sig(T, side="buy", atr=1.0, coin="ETH"):
    return {"symbol": coin, "signal": side, "mode": "X", "score": 1.0, "atr": atr, "T": T, "regime": "RANGING"}


class TestData(unittest.TestCase):
    def test_agg_bougies_completes_seulement(self):
        m1 = _m1(12)                                # 2 bougies M5 complètes + 2 M1 en trop
        c5 = B.agg(m1, 5)
        self.assertEqual(len(c5), 2)
        self.assertEqual(c5[0]["o"], m1[0]["o"])
        self.assertEqual(c5[0]["c"], m1[4]["c"])
        self.assertEqual(c5[0]["h"], max(x["h"] for x in m1[:5]))
        self.assertEqual(c5[0]["v"], 50.0)

    def test_window_exclut_la_bougie_en_cours(self):
        data = _data({"ETH": _m1(30)})
        w = B.window(data, "ETH", "5m", 12 * 60_000, 10)   # à 00:12, la M5 de 00:10 est en cours
        self.assertEqual([c["ts"] for c in w], [0, 300_000])
        w1 = B.window(data, "ETH", "1m", 12 * 60_000, 2)
        self.assertEqual(w1[-1]["ts"], 11 * 60_000)

    def test_regime_identique_au_bot(self):
        for btc, eth in ((10, 8), (-10, -9), (1, -2), (6, -6), (12, 10)):
            closes = {"BTC": [100.0] * 7 + [100 + btc], "ETH": [100.0] * 7 + [100 + eth]}
            if (btc, eth) == (12, 10):   # BULL en pullback 3 j → RANGING
                closes = {"BTC": [100, 100, 100, 100, 125, 120, 116, 112],
                          "ETH": [100, 100, 100, 100, 122, 118, 114, 110]}
            daily = {c: [(i * 86_400_000, float(px)) for i, px in enumerate(v)] for c, v in closes.items()}
            T_ = 8 * 86_400_000
            fc = lambda coin, iv, n, **k: [{"c": float(px)} for px in closes[coin]][-n:]
            with mock.patch.object(m, "fetch_candles", fc):
                bot = m.get_market_regime()[0]
            self.assertEqual(B.regime_at(daily, T_), bot, (btc, eth))


class TestSimulate(unittest.TestCase):
    def setUp(self):
        self.m1 = _m1(400)

    def sim(self, candles, sig, **k):
        return B.simulate(_data({"ETH": candles}), [sig], **k)

    def test_tp_achat_frais_deduits(self):
        cs = _m1(10) + [{"ts": 600_000, "o": 100, "h": 104.5, "l": 99.9, "c": 104, "v": 1}]
        tr = self.sim(cs, _sig(600_000), sl_mult=1.5, tp_mult=4.0, fee_pct=0.0864)
        self.assertEqual(tr[0]["how"], "TP")
        self.assertAlmostEqual(tr[0]["pnl"], 4.0 - 0.0864)

    def test_sl_prioritaire_meme_bougie(self):
        cs = _m1(10) + [{"ts": 600_000, "o": 100, "h": 105, "l": 98, "c": 100, "v": 1}]
        tr = self.sim(cs, _sig(600_000), fee_pct=0.0)
        self.assertEqual(tr[0]["how"], "SL")
        self.assertAlmostEqual(tr[0]["pnl"], -1.5)

    def test_short(self):
        cs = _m1(10) + [{"ts": 600_000, "o": 100, "h": 100.2, "l": 95.9, "c": 96, "v": 1}]
        tr = self.sim(cs, _sig(600_000, "sell"), fee_pct=0.0)
        self.assertEqual(tr[0]["how"], "TP")
        self.assertAlmostEqual(tr[0]["pnl"], 4.0)
        cs[-1] = {"ts": 600_000, "o": 100, "h": 101.6, "l": 99.9, "c": 101, "v": 1}
        self.assertEqual(self.sim(cs, _sig(600_000, "sell"), fee_pct=0.0)[0]["how"], "SL")

    def test_delai(self):
        tr = self.sim(_m1(500), _sig(0, atr=10.0), max_hours=1.0, fee_pct=0.0)
        self.assertEqual(tr[0]["how"], "Délai")
        self.assertAlmostEqual(tr[0]["dur_min"], 60)

    def test_une_position_par_coin(self):
        data = _data({"ETH": _m1(500)})
        tr = B.simulate(data, [_sig(0, atr=10.0), _sig(60_000, atr=10.0)], max_hours=1.0)
        self.assertEqual(len(tr), 1)

    def test_trailing_protege_le_gain(self):
        cs = _m1(5) + [{"ts": 300_000, "o": 100, "h": 102.0, "l": 99.9, "c": 101.9, "v": 1},
                       {"ts": 360_000, "o": 101.9, "h": 101.9, "l": 100.9, "c": 101, "v": 1}]
        tr = self.sim(cs, _sig(300_000), trail=True, fee_pct=0.0)
        self.assertEqual(tr[0]["how"], "BE")
        self.assertAlmostEqual(tr[0]["pnl"], 1.0)        # SL remonté à 102 − 1×ATR = 101


class TestStrategies(unittest.TestCase):
    def test_bot_strategy_reproduit_get_signal(self):
        fetch = lambda c, iv, n: T._declin(n)
        with mock.patch.object(m, "_current_regime", "BEAR"), \
             mock.patch.object(m, "get_penalty", lambda *a, **k: 0.0):
            s = B.bot_strategy(m)("ETH", fetch)
        self.assertEqual((s["signal"], s["mode"]), ("sell", "MOM"))
        self.assertGreater(s["atr"], 0)

    def test_scan_bout_en_bout(self):
        """Scan sur 20 h de bougies synthétiques avec une stratégie triviale."""
        data = _data({c: _m1(20 * 60, step=0.01) for c in B.COINS})
        daily = {c: [(i * 86_400_000, 100.0) for i in range(10)] for c in ("BTC", "ETH")}
        calls = []
        def strat(coin, fetch):
            calls.append(coin)
            return {"signal": "buy", "mode": "X", "score": 1.0, "atr": 1.0} if coin == "ETH" else None
        strat.STEP_MIN = 60
        sigs = B.scan(data, daily, strat)
        self.assertTrue(sigs)
        self.assertTrue(all(s["symbol"] == "ETH" and s["T"] % H == 0 for s in sigs))
        self.assertEqual(set(calls), set(B.COINS))

    def test_h1_trend_cassure_haussiere(self):
        cs = _m1(80 * 60, step=0.002)
        data = _data({"ETH": cs})
        Tn = cs[-1]["ts"] + 60_000
        s = S.h1_trend("ETH", lambda c, iv, n: B.window(data, c, iv, Tn, n))
        self.assertEqual(s["signal"], "buy")

    def test_rapport_contient_temoin(self):
        data = _data({c: _m1(3000) for c in B.COINS})
        sigs = [_sig(600_000 + i * 3_600_000, atr=0.5) for i in range(5)]
        txt = B.report(data, sigs, sl_mult=1.5, tp_mult=4.0, fee_pct=0.0864, max_hours=6.0)
        self.assertIn("BUY X", txt)
        self.assertIn("témoin", txt)


if __name__ == "__main__":
    unittest.main()
