"""
Tests de tools/trend_backtest.py — données synthétiques, sans réseau.
"""
import os, sys, time, unittest
from unittest import mock

import test_regression as T          # isolation de main.py (stubs, dossier temporaire) EN PREMIER
m = T.m
sys.path.insert(0, os.path.join(T.ROOT, "tools"))
import trend_backtest as B            # noqa: E402

D = 86_400_000


def _data(closes, lows=None, opens=None):
    """1 coin, une bougie par jour : haut = clôture + 0.5, bas = clôture − 0.5 (sauf lows), sans funding."""
    lows, opens = lows or {}, opens or {}
    ohlc = {i * D: (opens.get(i, c), c + 0.5, lows.get(i, c - 0.5), c) for i, c in enumerate(closes)}
    return {"ETH": {"days": sorted(ohlc), "ohlc": ohlc, "fund": {}, "fund_start": 0}}


class TestTrendBacktest(unittest.TestCase):
    RISE = [100.0 + i for i in range(20)]      # +1 par jour → TR = 1.5, ATR = 1.5

    def test_signal_long_si_cloture_au_dessus_de_celle_d_il_y_a_l_jours(self):
        sig = B.signals(_data([10, 12, 11, 9]), 2)["ETH"]
        self.assertEqual(sig, {2 * D: 1, 3 * D: -1})

    def test_position_tenue_jusqu_au_bout_frais_d_entree_compris(self):
        data = _data(self.RISE)
        curve = B.simulate(data, B.signals(data, 2), sl_atr=3.0, expo=0.5)
        q = 0.5 / 114                              # ouverture à la 15e clôture (ATR disponible)
        self.assertAlmostEqual(curve[-1][1], 1 - B.FEE * 0.5 + 5 * q, places=9)

    def test_sl_touche_sortie_au_prix_du_sl_puis_nouvelle_decision_a_la_cloture(self):
        closes = self.RISE + [112.0]
        data = _data(closes, lows={20: 105.0}, opens={20: 119.0})
        curve = B.simulate(data, B.signals(data, 2), sl_atr=3.0, expo=0.5)
        q, sl = 0.5 / 114, 114 - 3.0 * 1.5
        e_sl = 1 - B.FEE * 0.5 + 5 * q + q * (sl - 119) - B.FEE * q * sl
        q2 = 0.5 * e_sl / 112                      # short ouvert à la clôture du jour du SL (112 < 118)
        self.assertAlmostEqual(curve[-1][1], e_sl - B.FEE * q2 * 112, places=9)

    def test_meme_signal_que_le_bot(self):
        """Le backtest et main.trend_signal décident pareil sur les mêmes bougies."""
        closes = [100 + 7 * ((i * 37) % 11) - i for i in range(45)]
        y = (int(time.time()) // 86400 - 1) * D
        candles = [{"ts": y - (len(closes) - 1 - i) * D, "o": c, "h": c + 1, "l": c - 1, "c": c, "v": 1}
                   for i, c in enumerate(closes)]
        data = {"ETH": {"days": [c["ts"] for c in candles],
                        "ohlc": {c["ts"]: (c["o"], c["h"], c["l"], c["c"]) for c in candles},
                        "fund": {}, "fund_start": 0}}
        for end in range(31, 46):                  # 15 jours : les deux décisions coïncident chaque jour
            window = candles[:end]
            shift = (len(closes) - end) * D
            w = [dict(c, ts=c["ts"] + shift) for c in window]   # la dernière bougie = hier
            m._trend_sig.clear()
            with mock.patch.object(m, "fetch_candles", lambda *a, **k: w[-(m.TREND_LOOKBACK_D + 2):]):
                bot_side = m.trend_signal("ETH")[0]
            bt = B.signals(data, m.TREND_LOOKBACK_D)["ETH"][window[-1]["ts"]]
            self.assertEqual(bot_side, "BUY" if bt == 1 else "SELL")

    def test_temoin_decale_garde_les_memes_signaux(self):
        import random
        sig = {"ETH": {i * D: (1 if i % 3 else -1) for i in range(30)}}
        sh = B.shifted(sig, random.Random(1))["ETH"]
        self.assertEqual(sorted(sh.values()), sorted(sig["ETH"].values()))
        self.assertNotEqual(list(sh.values()), list(sig["ETH"].values()))


if __name__ == "__main__":
    unittest.main()
