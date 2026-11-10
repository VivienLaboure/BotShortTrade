"""
Tests de non-régression — BotShortTrade

Usage :
    python3 -m unittest discover -s tests -v

Aucun appel réseau, aucun ordre réel, aucun message Discord :
  - dotenv neutralisé (le vrai .env n'est jamais lu) + clé privée factice
  - SDK Hyperliquid remplacé par des stubs avant l'import de main.py
  - requests.get/post lèvent une erreur si un test oublie de les mocker
  - exécution dans un dossier temporaire (trades.xlsx, bot_state.json… factices)
"""
import json, os, sys, tempfile, time, types, unittest
from datetime import datetime
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Isolation AVANT l'import de main.py ──────────────────────────────────────
_dotenv = types.ModuleType("dotenv")
_dotenv.load_dotenv = lambda *a, **k: None
sys.modules["dotenv"] = _dotenv

os.environ.update({
    "HL_PRIVATE_KEY": "0x" + "11" * 32,   # clé factice
    "HL_TESTNET": "1",
    "DISCORD_WEBHOOK_URL": "",
    "LEVERAGE": "5", "CAPITAL_PCT": "0.15", "VOL_MULTIPLIER": "1.2",
    "MIN_MOVE_PCT": "0.0002", "MAX_TRADES": "3", "HEARTBEAT_HOURS": "6",
})

class _StubInfo:
    def __init__(self, *a, **k): self.kwargs = k
class _StubExchange:
    def __init__(self, *a, **k): self.kwargs = k
for name, attr, cls in [("hyperliquid.info", "Info", _StubInfo),
                        ("hyperliquid.exchange", "Exchange", _StubExchange)]:
    mod = types.ModuleType(name); setattr(mod, attr, cls); sys.modules[name] = mod
if "hyperliquid" not in sys.modules:
    sys.modules["hyperliquid"] = types.ModuleType("hyperliquid")
try:
    import hyperliquid.utils.constants  # noqa: F401  (vrai module, sans réseau)
except Exception:
    _u = types.ModuleType("hyperliquid.utils"); _c = types.ModuleType("hyperliquid.utils.constants")
    _c.TESTNET_API_URL = _c.MAINNET_API_URL = "http://stub"; _u.constants = _c
    sys.modules["hyperliquid.utils"] = _u; sys.modules["hyperliquid.utils.constants"] = _c

import requests
def _no_network(*a, **k):
    raise AssertionError(f"Appel réseau interdit dans les tests : {a[:1]}")
requests.get = requests.post = _no_network

_TMP = tempfile.mkdtemp(prefix="bst_tests_")
os.chdir(_TMP)
sys.path.insert(0, ROOT)
import contextlib, io
with contextlib.redirect_stdout(io.StringIO()):
    import main as m
import openpyxl


# ── Helpers ──────────────────────────────────────────────────────────────────
def _ok(oid=2):
    return {"status": "ok", "response": {"data": {"statuses": [{"resting": {"oid": oid}}]}}}

def _err(msg="Invalid TP/SL price"):
    return {"status": "ok", "response": {"data": {"statuses": [{"error": msg}]}}}

class FakeExchange:
    """Simule Hyperliquid : fill à avgPx, réponses SL/TP/close configurables."""
    def __init__(self, avg_px=2010.0, sl=None, tp=None, close_ok=True):
        self.avg_px, self.close_ok = avg_px, close_ok
        self.sl = sl or [_ok(11)]
        self.tp = tp or [_ok(12)]
        self.orders, self.closed = [], 0
    def update_leverage(self, *a, **k): return {"status": "ok"}
    def market_open(self, *a, **k):
        return {"status": "ok", "response": {"data": {"statuses": [
            {"filled": {"oid": 1, "avgPx": str(self.avg_px), "totalSz": "0.04"}}]}}}
    def order(self, coin, is_buy, qty, limit_px, order_type, reduce_only):
        trig = order_type["trigger"]; self.orders.append(trig)
        seq = self.sl if trig["tpsl"] == "sl" else self.tp
        return seq.pop(0) if len(seq) > 1 else seq[0]
    def market_close(self, coin, slippage=0.05):
        self.closed += 1
        return {"status": "ok"} if self.close_ok else {"status": "err"}
    def cancel(self, *a, **k): return {"status": "ok"}

class FakeInfo:
    def __init__(self, mids=None, perp=None, spot=None, fills=None, abstraction="unifiedAccount"):
        self.mids = mids or {"ETH": "2000"}
        self.perp, self.spot, self.fills = perp or {}, spot or {"balances": []}, fills or []
        self.funding, self.rates = [], []   # paiements de financement du compte, taux horaires publics
        self.ledger = []                     # journal des dépôts, retraits et virements
        self.abstraction = abstraction       # mode du compte (celui du bot : Unified)
    def all_mids(self): return self.mids
    def user_state(self, addr): return self.perp
    def spot_user_state(self, addr): return self.spot
    def query_user_abstraction_state(self, addr): return self.abstraction
    def user_fills(self, addr): return self.fills
    def user_funding_history(self, addr, start): return [f for f in self.funding if f["time"] >= start]
    def funding_history(self, coin, start): return [r for r in self.rates if r["time"] >= start]
    def user_non_funding_ledger_updates(self, addr, start): return [r for r in self.ledger if r["time"] >= start]

SIG = {"symbol": "ETH", "signal": "buy", "score": 0.8, "atr_m5": 5.0, "price": 2000,
       "conditions": {"mode": "MR", "regime": "RANGING", "rsi": 50, "bbp": 0.3, "vol_ratio": 1.5}}


class BaseTest(unittest.TestCase):
    def setUp(self):
        for f in (m.EXCEL_FILE, m.STATE_FILE, m.LESSONS_FILE, m.TREND_FILE):
            if os.path.exists(f):
                os.remove(f)
        m._trend = None; m._trend_sig.clear(); m._trend_alerted.clear(); m._trend_sl_ts = 0.0
        m._unified = None
        self.alerts = []
        self._p = [
            mock.patch.object(m, "send_discord_alert",
                              lambda t, d, color=0: self.alerts.append((t, d))),
            mock.patch.object(m.time, "sleep", lambda s: None),
            mock.patch("sys.stdout", new_callable=io.StringIO),
        ]
        for p in self._p: p.start()
    def tearDown(self):
        for p in self._p: p.stop()
    def titles(self):
        return [t for t, _ in self.alerts]


def _hl_px_valide(px: float, sz_dec: int = 0) -> bool:
    """Règle Hyperliquid (perps) : entier, ou ≤ 5 chiffres significatifs et ≤ 6 − szDecimals décimales."""
    if px == int(px):
        return True
    txt = f"{px:.10f}".rstrip("0")
    return len(txt.replace(".", "").lstrip("0")) <= 5 and len(txt.split(".")[1]) <= 6 - sz_dec


# ── place_order : SL/TP et sécurité stop-loss ────────────────────────────────
class TestPlaceOrder(BaseTest):
    def run_order(self, ex, sig=SIG):
        with mock.patch.object(m, "exchange", ex), mock.patch.object(m, "info", FakeInfo()):
            wb = m.init_excel()
            return m.place_order(dict(sig), wb), wb

    def test_sl_tp_calcules_sur_prix_execute(self):
        pos, _ = self.run_order(FakeExchange(avg_px=2010.0))
        self.assertEqual(pos["entry"], 2010.0)
        self.assertAlmostEqual(pos["sl"], 2010.0 - m.ATR_SL_MULT * 5.0)
        self.assertAlmostEqual(pos["tp"], 2010.0 + m.ATR_TP_MULT * 5.0)
        self.assertEqual(pos["sl_oid"], 11)

    def test_sl_tp_short(self):
        sig = dict(SIG, signal="sell")
        pos, _ = self.run_order(FakeExchange(avg_px=1990.0), sig)
        self.assertAlmostEqual(pos["sl"], 1990.0 + m.ATR_SL_MULT * 5.0)
        self.assertAlmostEqual(pos["tp"], 1990.0 - m.ATR_TP_MULT * 5.0)

    def test_prix_sl_tp_valides_doge_au_dessus_de_10_cents(self):
        """DOGE à 0.12345 : l'ancien arrondi (6 décimales) donnait 6 chiffres significatifs,
        refusés par Hyperliquid → SL refusé → fermeture d'urgence de chaque trade."""
        sig = dict(SIG, symbol="DOGE", atr_m5=0.000123, price=0.12345)
        ex = FakeExchange(avg_px=0.12345)
        with mock.patch.object(m, "exchange", ex), \
             mock.patch.object(m, "info", FakeInfo(mids={"DOGE": "0.12345"})):
            pos = m.place_order(sig, m.init_excel())
        self.assertEqual(ex.closed, 0)
        for o in ex.orders:
            self.assertTrue(_hl_px_valide(o["triggerPx"]), o)
        self.assertTrue(_hl_px_valide(pos["sl"]) and _hl_px_valide(pos["tp"]))

    def test_sl_refuse_puis_accepte_au_retry(self):
        ex = FakeExchange(sl=[_err(), _ok(33)])
        pos, _ = self.run_order(ex)
        self.assertEqual(pos["sl_oid"], 33)
        self.assertEqual(ex.closed, 0)

    def test_sl_refuse_deux_fois_ferme_la_position(self):
        ex = FakeExchange(sl=[_err()])
        with self.assertRaises(RuntimeError):
            self.run_order(ex)
        self.assertEqual(ex.closed, 1)
        self.assertTrue(any("SL refusé" in t for t in self.titles()))
        self.assertFalse(any(o["tpsl"] == "tp" for o in ex.orders))   # pas de TP posé

    def test_sl_status_error_global(self):
        ex = FakeExchange(sl=[{"status": "err", "response": "rate limited"}])
        with self.assertRaises(RuntimeError):
            self.run_order(ex)
        self.assertEqual(ex.closed, 1)

    def test_sl_refuse_et_fermeture_echoue_alerte_critique(self):
        ex = FakeExchange(sl=[_err()], close_ok=False)
        pos, _ = self.run_order(ex)          # position renvoyée pour rester suivie
        self.assertIsNone(pos["sl_oid"])
        self.assertTrue(any("SANS STOP-LOSS" in t for t in self.titles()))

    def test_tp_refuse_alerte_mais_position_gardee(self):
        ex = FakeExchange(tp=[_err()])
        pos, _ = self.run_order(ex)
        self.assertEqual(pos["sl_oid"], 11)
        self.assertEqual(ex.closed, 0)
        self.assertTrue(any("TP refusé" in t for t in self.titles()))

    def test_ligne_excel_mode_regime_score(self):
        _, wb = self.run_order(FakeExchange())
        ws = wb["Trades"]
        row = [c.value for c in ws[ws.max_row]]
        hdr = [c.value for c in ws[1]]
        self.assertEqual(row[hdr.index("Mode")], "MR")
        self.assertEqual(row[hdr.index("Régime")], "RANGING")
        self.assertEqual(row[hdr.index("Score")], 0.8)
        self.assertEqual(hdr[14], "P&L final $")   # index utilisé par analyze_weekly.py


# ── Excel ────────────────────────────────────────────────────────────────────
class TestExcel(BaseTest):
    def test_migration_ancien_fichier_15_colonnes(self):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Trades"
        ws.append(m.COL_HEADERS[:15]); ws.append(["1", "2026-01-01 00:00"] + [0] * 13)
        wb.save(m.EXCEL_FILE)
        ws2 = m.init_excel()["Trades"]
        self.assertEqual([c.value for c in ws2[1]], m.COL_HEADERS)
        self.assertEqual(ws2.cell(2, 1).value, "1")   # données conservées

    def test_fichier_illisible_garde_sous_un_autre_nom(self):
        """trades.xlsx illisible (abîmé, réenregistré dans un autre format) : gardé sous un autre nom avec
        une alerte, puis nouveau fichier (avant : écrasé par un classeur vide, historique perdu)."""
        with open(m.EXCEL_FILE, "wb") as f:
            f.write(b"pas un classeur")
        wb = m.init_excel()
        kept = [f for f in os.listdir(".") if f.startswith("trades-illisible-")]
        try:
            self.assertEqual(len(kept), 1)
            self.assertTrue(kept[0].endswith(".xlsx"))         # ignoré par git (*.xlsx)
            with open(kept[0], "rb") as f:
                self.assertEqual(f.read(), b"pas un classeur")
        finally:
            for f in kept:
                os.remove(f)
        self.assertIn("Tendance", wb.sheetnames)
        self.assertIn("⚠️ trades.xlsx illisible", self.titles())


# ── Équité ───────────────────────────────────────────────────────────────────
class TestEquity(BaseTest):
    def eq(self, perp, spot, abstraction="unifiedAccount"):
        with mock.patch.object(m, "info", FakeInfo(spot=spot, abstraction=abstraction)):
            return m.get_equity(perp)

    def test_unified_marge_pas_comptee_deux_fois(self):
        # Valeurs réelles relevées le 2026-10-16 avec 1 position DOGE ouverte : le total spot contient la marge
        perp = {"withdrawable": "0.0", "crossMarginSummary":
                {"accountValue": "16.556911", "totalMarginUsed": "16.68719"}}
        spot = {"balances": [{"coin": "USDC", "total": "111.343455", "hold": "16.68719"}]}
        self.assertAlmostEqual(self.eq(perp, spot), 111.343455, places=6)

    def test_unified_pnl_latent_pas_compte_deux_fois(self):
        """06/11, 5 shorts ouverts, P&L latent +0.41 $ : Hyperliquid (endpoint portfolio) donnait 211.3697 $,
        le total spot ; le bot affichait 211.87 $ (P&L latent compté environ 2,2 fois)."""
        perp = {"withdrawable": "0.497491", "crossMarginSummary":
                {"accountValue": "21.880433", "totalNtlPos": "106.88072", "totalRawUsd": "128.761153",
                 "totalMarginUsed": "21.376143"}}
        spot = {"balances": [{"coin": "USDC", "total": "211.3697", "hold": "21.376143"}]}
        self.assertAlmostEqual(self.eq(perp, spot), 211.3697, places=6)

    def test_mode_du_compte_lu_une_fois(self):
        info = FakeInfo(spot={"balances": [{"coin": "USDC", "total": "100", "hold": "0"}]})
        info.query_user_abstraction_state = mock.Mock(side_effect=[Exception("(502, 'Bad Gateway')"),
                                                                   "unifiedAccount"])
        perp = {"crossMarginSummary": {"accountValue": "0.0", "totalMarginUsed": "0.0"}}
        with mock.patch.object(m, "info", info):
            with self.assertRaises(Exception):          # l'appelant garde le solde précédent
                m.get_equity(perp)
            self.assertEqual([m.get_equity(perp) for _ in range(3)], [100.0] * 3)
        self.assertEqual(info.query_user_abstraction_state.call_count, 2)

    def test_unified_sans_position(self):
        perp = {"withdrawable": "0.0", "crossMarginSummary":
                {"accountValue": "0.0", "totalMarginUsed": "0.0"}}
        spot = {"balances": [{"coin": "USDC", "total": "111.5483", "hold": "0.0"}]}
        self.assertAlmostEqual(self.eq(perp, spot), 111.5483, places=4)

    def test_compte_classique_perp_seul(self):
        perp = {"withdrawable": "80", "crossMarginSummary":
                {"accountValue": "100", "totalMarginUsed": "20"}}
        self.assertAlmostEqual(self.eq(perp, {"balances": []}, "default"), 100.0)

    def test_compte_classique_perp_plus_spot(self):
        perp = {"crossMarginSummary": {"accountValue": "100", "totalMarginUsed": "20"}}
        spot = {"balances": [{"coin": "USDC", "total": "50", "hold": "0"}]}
        self.assertAlmostEqual(self.eq(perp, spot, "default"), 150.0)

    def test_spot_indisponible_leve_au_lieu_de_zero(self):
        """25/10 : 504 sur spot_user_state → équité lue $0.00 → faux flash crash -100 %."""
        info = FakeInfo(spot={})
        info.spot_user_state = mock.Mock(side_effect=Exception("(504, 'Gateway Timeout')"))
        perp = {"crossMarginSummary": {"accountValue": "0.0", "totalMarginUsed": "0.0"}}
        with mock.patch.object(m, "info", info), self.assertRaises(Exception):
            m.get_equity(perp)

    def test_user_state_vide_leve(self):
        with mock.patch.object(m, "info", FakeInfo(spot={"balances": []})), self.assertRaises(ValueError):
            m.get_equity({})


# ── P&L : frais d'ouverture ──────────────────────────────────────────────────
class TestSharedWallet(BaseTest):
    POS = [{"position": {"coin": "ETH", "szi": "0.01", "entryPx": "2000", "unrealizedPnl": "1.5"}},
           {"position": {"coin": "SOL", "szi": "-0.5", "entryPx": "100", "unrealizedPnl": "-3"}},
           {"position": {"coin": "XRP", "szi": "0", "entryPx": "1", "unrealizedPnl": "0"}}]

    def test_account_coins(self):
        self.assertEqual(m.account_coins({"assetPositions": self.POS}), {"ETH", "SOL"})
        self.assertEqual(m.account_coins(None), set())

    def test_pnl_latent_limite_a_nos_positions(self):
        st = {"assetPositions": self.POS}
        self.assertAlmostEqual(m.get_unrealized_pnl(st, {"ETH"}), 1.5)
        self.assertAlmostEqual(m.get_unrealized_pnl(st), -1.5)

    def test_reprise_n_adopte_que_nos_positions(self):
        m.save_positions_state([{"id": "o1", "symbol": "ETH", "side": "BUY", "entry": 2000,
                                 "qty": 0.01, "sl": 1990, "tp": 2030, "sl_oid": 7, "atr": 5,
                                 "open_fee": 0.01, "opened_at": 1.0, "conditions": {}}])
        with mock.patch.object(m, "info", FakeInfo(perp={"assetPositions": self.POS})):
            rec, _ = m.recover_open_positions()
        self.assertEqual([p["symbol"] for p in rec], ["ETH"])
        self.assertEqual(rec[0]["sl_oid"], 7)
        self.assertEqual(self.titles(), ["⚠️ Position non reprise"])


class TestFees(BaseTest):
    def test_zombie_pnl_deduit_frais_ouverture_estimes(self):
        fills = [{"coin": "ETH", "dir": "Close Long", "time": 1, "fee": "0.04", "closedPnl": "1.00"}]
        pos = {"id": "1", "symbol": "ETH", "side": "BUY", "entry": 2000, "qty": 0.04,
               "sl_oid": None, "open_fee": 0.0, "conditions": {}}
        with mock.patch.object(m, "exchange", FakeExchange()), \
             mock.patch.object(m, "info", FakeInfo(fills=fills)), \
             mock.patch.object(m, "update_order_status", lambda *a, **k: None), \
             mock.patch.object(m, "record_lesson", lambda *a, **k: None):
            pnl = m.force_close_zombie(m.init_excel(), pos)
        self.assertAlmostEqual(pnl, 1.00 - 0.04 - 0.04, places=2)


    # Fills réels du trade DOGE du 2026-10-16 : SL exécuté en 3 fills partiels.
    # Avant correctif : seul le dernier fill était compté → P&L −0.08 au lieu de −0.37.
    DOGE_FILLS = [
        {"coin": "DOGE", "dir": "Open Long",  "time": 1792136736618, "oid": 576377349716,
         "sz": "1016.0", "closedPnl": "0.0", "fee": "0.036116"},
        {"coin": "DOGE", "dir": "Close Long", "time": 1792138739787, "oid": 576377359631,
         "sz": "222.0", "closedPnl": "-0.0666", "fee": "0.007862"},
        {"coin": "DOGE", "dir": "Close Long", "time": 1792138739787, "oid": 576377359631,
         "sz": "677.0", "closedPnl": "-0.195653", "fee": "0.023981"},
        {"coin": "DOGE", "dir": "Close Long", "time": 1792138739787, "oid": 576377359631,
         "sz": "117.0", "closedPnl": "-0.033579", "fee": "0.004144"},
        # Ancien trade DOGE fermé plus tôt (autre oid) : ne doit PAS être additionné
        {"coin": "DOGE", "dir": "Close Short", "time": 1792000000000, "oid": 111,
         "sz": "500.0", "closedPnl": "5.0", "fee": "0.02"},
    ]
    DOGE_POS = {"id": "576377349716", "symbol": "DOGE", "side": "BUY", "entry": 0.082287,
                "qty": 1016.0, "sl": 0.082014, "tp": 0.083016, "sl_oid": 576377359631,
                "atr": 0.000182, "open_fee": 0.0, "conditions": {"mode": "MR"}}

    def test_sl_en_fills_partiels_tous_additionnes(self):
        rec = {}
        with mock.patch.object(m, "info", FakeInfo(fills=self.DOGE_FILLS)), \
             mock.patch.object(m, "update_order_status",
                               lambda wb, oid, status, pnl=None: rec.update(status=status, pnl=pnl)), \
             mock.patch.object(m, "record_lesson", lambda *a, **k: None):
            result = m.check_position(None, dict(self.DOGE_POS), cached_state={"assetPositions": []})
        self.assertEqual(result, "SL touché")
        # gross −0.295832 − frais ouverture réels 0.036116 − frais sortie 0.035987
        self.assertAlmostEqual(rec["pnl"], -0.37, places=2)

    def test_closing_pnl_frais_ouverture_reels(self):
        gross, open_fee, close_fee = m._closing_pnl(self.DOGE_POS, self.DOGE_FILLS)
        self.assertAlmostEqual(gross, -0.295832, places=6)
        self.assertAlmostEqual(open_fee, 0.036116, places=6)
        self.assertAlmostEqual(close_fee, 0.035987, places=6)


# ── Revue 2026-10-16 : zombie, attribution des fermetures, surveillance SL ───
class OrdersInfo(FakeInfo):
    def __init__(self, orders=None, **k):
        super().__init__(**k); self.orders = orders or []
    def frontend_open_orders(self, addr): return self.orders

class LogExchange(FakeExchange):
    def __init__(self, close_resp=None, **k):
        super().__init__(**k); self.log, self.close_resp = [], close_resp
    def market_close(self, coin, slippage=0.05):
        self.log.append("close"); return self.close_resp
    def cancel(self, coin, oid):
        self.log.append(("cancel", oid)); return {"status": "ok"}
    def order(self, *a, **k):
        self.log.append("order"); return super().order(*a, **k)

class TestZombie(BaseTest):
    POS = {"id": "1", "symbol": "ETH", "side": "BUY", "entry": 2000, "qty": 0.04,
           "sl_oid": 11, "open_fee": 0.0, "conditions": {}}

    def close(self, resp):
        ex = LogExchange(close_resp=resp)
        with mock.patch.object(m, "exchange", ex), mock.patch.object(m, "info", FakeInfo()), \
             mock.patch.object(m, "update_order_status", lambda *a, **k: None), \
             mock.patch.object(m, "record_lesson", lambda *a, **k: None):
            return m.force_close_zombie(None, dict(self.POS)), ex

    def test_fermeture_puis_annulation_sl(self):
        pnl, ex = self.close({"status": "ok", "response": {"data": {"statuses": [{"filled": {}}]}}})
        self.assertIsNotNone(pnl)
        self.assertEqual(ex.log, ["close", ("cancel", 11)])

    def test_fermeture_echouee_sl_conserve(self):
        pnl, ex = self.close({"status": "err", "response": "insufficient liquidity"})
        self.assertIsNone(pnl)
        self.assertEqual(ex.log, ["close"])                  # SL jamais annulé
        self.assertTrue(any("fermeture échouée" in t for t in self.titles()))

    def test_fermeture_statut_erreur_sl_conserve(self):
        pnl, ex = self.close({"status": "ok", "response": {"data": {"statuses": [{"error": "x"}]}}})
        self.assertIsNone(pnl)
        self.assertEqual(ex.log, ["close"])


class TestClosingAttribution(BaseTest):
    def test_fermeture_anterieure_a_l_ouverture_ignoree(self):
        fills = [{"coin": "DOGE", "dir": "Close Short", "time": 1_000_000, "oid": 9,
                  "closedPnl": "5.0", "fee": "0.02"}]
        pos = {"id": "1", "symbol": "DOGE", "opened_at": 2_000.0}   # ouvert après ce fill
        self.assertIsNone(m._closing_pnl(pos, fills))

    def test_position_absente_sans_fill_attend_3_cycles(self):
        pos = {"id": "1", "symbol": "DOGE", "side": "BUY", "entry": 0.08, "qty": 1000,
               "opened_at": 2_000_000_000.0, "conditions": {}}
        rec = {}
        with mock.patch.object(m, "info", FakeInfo(fills=[])), \
             mock.patch.object(m, "update_order_status",
                               lambda wb, oid, status, pnl=None: rec.update(status=status)), \
             mock.patch.object(m, "record_lesson", lambda *a, **k: None):
            r1 = m.check_position(None, pos, cached_state={"assetPositions": []})
            r2 = m.check_position(None, pos, cached_state={"assetPositions": []})
            r3 = m.check_position(None, pos, cached_state={"assetPositions": []})
        self.assertEqual((r1, r2), (None, None))
        self.assertEqual(r3, "Fermé (inconnu)")
        self.assertEqual(rec["status"], "Fermé (inconnu)")


class TestSLWatchdog(BaseTest):
    POS = {"id": "1", "symbol": "ETH", "side": "BUY", "entry": 2000.0, "qty": 0.04,
           "sl": 1992.5, "sl_oid": None}
    STOP = {"coin": "ETH", "isTrigger": True, "reduceOnly": True, "orderType": "Stop Market", "oid": 55}
    TP   = {"coin": "ETH", "isTrigger": True, "reduceOnly": True, "orderType": "Take Profit Market", "oid": 56}

    def watch(self, orders, ex=None):
        pos, ex = dict(self.POS), ex or LogExchange()
        with mock.patch.object(m, "info", OrdersInfo(orders=orders)), mock.patch.object(m, "exchange", ex):
            n = m.ensure_sl_orders([pos])
        return n, pos, ex

    def test_sl_present_resynchronise_oid(self):
        n, pos, ex = self.watch([self.STOP, self.TP])
        self.assertEqual(n, 0)
        self.assertEqual(pos["sl_oid"], 55)
        self.assertEqual(ex.log, [])

    def test_sl_absent_repose(self):
        n, pos, ex = self.watch([self.TP])                 # seul le TP existe
        self.assertEqual(n, 1)
        self.assertEqual(ex.orders[0]["triggerPx"], 1992.5)
        self.assertEqual(ex.orders[0]["tpsl"], "sl")
        self.assertEqual(pos["sl_oid"], 11)
        self.assertTrue(any("SL manquant" in t for t in self.titles()))

    def test_stop_du_meme_sens_que_la_position_ignore(self):
        """Un stop d'achat ne protège pas un long (SL orphelin d'un ancien short) : SL reposé (avant : pris
        pour le SL du long, qui restait sans protection)."""
        n, pos, ex = self.watch([dict(self.STOP, side="B")])
        self.assertEqual(n, 1)
        self.assertEqual(pos["sl_oid"], 11)
        n, pos, ex = self.watch([dict(self.STOP, side="A")])    # stop de vente : c'est bien son SL
        self.assertEqual((n, pos["sl_oid"]), (0, 55))

    def test_sl_absent_et_refuse_alerte_critique(self):
        n, pos, ex = self.watch([], LogExchange(sl=[_err()]))
        self.assertEqual(n, 0)
        self.assertTrue(any("SANS STOP-LOSS" in t for t in self.titles()))


class TestDailyPnl(BaseTest):
    def test_pnl_du_jour_par_date_de_fermeture(self):
        wb = m.init_excel(); ws = wb["Trades"]
        today = m.datetime.now(m.timezone.utc).strftime("%Y-%m-%d")
        base = ["x"] * 13
        ws.append(["1", "2000-01-01 23:50"] + base[:11] + ["SL touché", -1.0, "MR", "RANGING", 0.8, f"{today} 00:10"])
        ws.append(["2", f"{today} 08:00"] + base[:11] + ["TP touché", 0.5])          # ancienne ligne sans date
        ws.append(["3", f"{today} 09:00"] + base[:11] + ["SL touché", -0.3, "MR", "RANGING", 0.8, "2000-01-02 00:00"])
        self.assertAlmostEqual(m.calc_daily_pnl(wb), -0.5)

    def test_date_fermeture_ecrite(self):
        with mock.patch.object(m, "exchange", FakeExchange()), mock.patch.object(m, "info", FakeInfo()):
            wb = m.init_excel(); m.place_order(dict(SIG), wb)
        oid = wb["Trades"].cell(2, 1).value
        m.update_order_status(wb, oid, "TP touché", 0.7)
        self.assertTrue(str(wb["Trades"].cell(2, 19).value).startswith(
            m.datetime.now(m.timezone.utc).strftime("%Y-%m-%d")))


# ── Trailing stop ────────────────────────────────────────────────────────────
class TrailExchange(FakeExchange):
    def __init__(self, resp):
        super().__init__(); self.resp, self.log = resp, []
    def order(self, *a, **k):
        self.log.append("order"); return self.resp
    def cancel(self, coin, oid):
        self.log.append(("cancel", oid)); return {"status": "ok"}

class TestTrailing(BaseTest):
    POS = {"symbol": "ETH", "side": "BUY", "entry": 2000.0, "qty": 0.04,
           "sl": 1992.5, "sl_oid": 11, "atr": 5.0}

    def trail(self, resp, price=2010.0):   # +2×ATR → phase trailing
        pos, ex = dict(self.POS), TrailExchange(resp)
        with mock.patch.object(m, "exchange", ex):
            m.trail_sl(pos, price)
        return pos, ex

    def test_nouveau_sl_pose_avant_annulation_ancien(self):
        pos, ex = self.trail(_ok(77))
        self.assertEqual(ex.log, ["order", ("cancel", 11)])
        self.assertEqual(pos["sl_oid"], 77)
        self.assertAlmostEqual(pos["sl"], 2010.0 - m.TRAIL_DIST * 5.0)
        self.assertTrue(pos["trailed"])

    def test_nouveau_sl_refuse_ancien_conserve(self):
        pos, ex = self.trail(_err())
        self.assertEqual(ex.log, ["order"])            # ancien SL jamais annulé
        self.assertEqual(pos["sl_oid"], 11)
        self.assertEqual(pos["sl"], 1992.5)
        self.assertNotIn("trailed", pos)

    def test_nouveau_sl_status_erreur_ancien_conserve(self):
        pos, ex = self.trail({"status": "err", "response": "rate limited"})
        self.assertEqual(ex.log, ["order"])
        self.assertEqual(pos["sl_oid"], 11)

    def test_breakeven_couvre_les_frais(self):
        pos, ex = self.trail(_ok(77), price=2006.0)   # +1.2×ATR → phase breakeven
        self.assertAlmostEqual(pos["sl"], 2000.0 * (1 + m.BE_BUFFER_PCT / 100))

    def test_breakeven_short(self):
        pos = dict(self.POS, side="SELL", sl=2007.5)
        with mock.patch.object(m, "exchange", TrailExchange(_ok(78))):
            m.trail_sl(pos, 1994.0)
        self.assertAlmostEqual(pos["sl"], 2000.0 * (1 - m.BE_BUFFER_PCT / 100))

    def test_pas_de_trail_sous_1_atr(self):
        pos, ex = self.trail(_ok(77), price=2003.0)
        self.assertEqual(ex.log, [])


# ── Apprentissage ────────────────────────────────────────────────────────────
class TestLessons(BaseTest):
    POS = {"symbol": "DOGE", "side": "BUY",
           "conditions": {"rsi": 52.9, "bbp": 0.346, "regime": "RANGING", "mode": "MR"}}

    def setUp(self):
        super().setUp()
        m._lessons_cache.clear(); m._lessons_cache_ts = 0.0

    def test_trail_ferme_a_perte_compte_comme_perte(self):
        m.record_lesson(self.POS, "TRAIL touché", -0.05)
        b = m._load_lessons()["buckets"]["coin_DOGE|buy"]
        self.assertEqual((b["w"], b["l"]), (0, 1))

    def test_trail_ferme_en_gain_compte_comme_gain(self):
        m.record_lesson(self.POS, "TRAIL touché", 0.40)
        b = m._load_lessons()["buckets"]["coin_DOGE|buy"]
        self.assertEqual((b["w"], b["l"]), (1, 0))

    def pen(self):
        return m.get_penalty("DOGE", "BUY", 52.9, 0.346, "RANGING", "MR")

    def test_penalite_apres_6_pertes_avec_oubli(self):
        for _ in range(5):
            m.record_lesson(self.POS, "SL touché", -0.37)
        self.assertEqual(self.pen(), 0.0)   # oubli : 5 pertes ≈ 4.71 effectives < 5
        m.record_lesson(self.POS, "SL touché", -0.37)
        self.assertAlmostEqual(self.pen(), 0.25)   # 4 buckets à 0 % WR → plafond

    def test_bonus_si_historique_gagnant(self):
        for _ in range(8):
            m.record_lesson(self.POS, "TP touché", 0.70)
        self.assertAlmostEqual(self.pen(), -0.15)  # bonus plafonné à 0.15

    def test_oubli_progressif(self):
        m.record_lesson(self.POS, "SL touché", -0.37)
        m.record_lesson(self.POS, "TP touché", 0.70)
        b = m._load_lessons()["buckets"]["coin_DOGE|buy"]
        self.assertAlmostEqual(b["l"], 0.97)
        self.assertAlmostEqual(b["w"], 1.0)

    def test_fantomes_comptent_a_moitie(self):
        for _ in range(10):
            m.record_lesson(self.POS, "SL", -0.5, ghost=True)
        self.assertEqual(self.pen(), 0.0)          # 10 fantômes ≈ 4.3 effectifs
        for _ in range(2):
            m.record_lesson(self.POS, "SL", -0.5, ghost=True)
        self.assertGreater(self.pen(), 0.0)
        b = m._load_lessons()["buckets"]["coin_DOGE|buy"]
        self.assertEqual((b["w"], b["l"]), (0, 0))  # compteurs réels intacts
        self.assertAlmostEqual(b["gpnl_pct"], -6.0)


# ── Mode fantôme ─────────────────────────────────────────────────────────────
def _m1(ts, o, h, l, c):
    return {"ts": ts, "o": o, "h": h, "l": l, "c": c, "v": 1.0}

class TestGhosts(BaseTest):
    def setUp(self):
        super().setUp()
        if os.path.exists(m.GHOST_FILE):
            os.remove(m.GHOST_FILE)
        m._ghosts = None
        m._lessons_cache.clear(); m._lessons_cache_ts = 0.0
        self.wb = m.init_excel()

    def g(self, side="buy"):
        return {"side": side, "entry": 100.0, "sl": 99.0 if side == "buy" else 101.0,
                "tp": 102.0 if side == "buy" else 98.0, "checked_ts": 0, "opened_at": 1000.0}

    def test_tp_touche(self):
        g = self.g()
        self.assertEqual(m._ghost_resolve(g, [_m1(1, 100, 101, 99.5, 101), _m1(2, 101, 102.1, 100.5, 102)], 1100),
                         ("TP", 102.0))
        self.assertEqual(g["checked_ts"], 2)

    def test_sl_et_tp_meme_bougie_compte_sl(self):
        self.assertEqual(m._ghost_resolve(self.g(), [_m1(1, 100, 102.5, 98.5, 100)], 1100), ("SL", 99.0))

    def test_short(self):
        self.assertEqual(m._ghost_resolve(self.g("sell"), [_m1(1, 100, 100.2, 97.9, 98)], 1100), ("TP", 98.0))
        self.assertEqual(m._ghost_resolve(self.g("sell"), [_m1(1, 100, 101.2, 99.9, 101)], 1100), ("SL", 101.0))

    def test_bougies_deja_vues_ignorees(self):
        g = dict(self.g(), checked_ts=5)
        self.assertIsNone(m._ghost_resolve(g, [_m1(5, 100, 103, 98, 100)], 1100))

    def test_delai_depasse(self):
        g = self.g()
        res = m._ghost_resolve(g, [_m1(1, 100, 100.5, 99.5, 100.4)], 1000 + m.GHOST_MAX_HOURS * 3600)
        self.assertEqual(res, ("Délai", 100.4))

    def test_pnl_pct_net_de_frais(self):
        self.assertAlmostEqual(m._ghost_pnl_pct(self.g(), 102.0), 2.0 - m.GHOST_FEE_PCT)
        self.assertAlmostEqual(m._ghost_pnl_pct(self.g("sell"), 101.0), -1.0 - m.GHOST_FEE_PCT)

    def test_ouverture_dedoublonnee_et_excel(self):
        sig = dict(SIG, price=100.0, atr_m5=1.0)
        self.assertIsNotNone(m.ghost_open(self.wb, sig, "score<0.75"))
        self.assertIsNone(m.ghost_open(self.wb, sig, "score<0.75"))   # même coin/côté/mode
        ws = self.wb["Fantômes"]
        self.assertEqual(ws.max_row, 2)
        self.assertEqual(ws.cell(2, 8).value, "score<0.75")
        self.assertTrue(os.path.exists(m.GHOST_FILE))

    def test_cycle_complet_tp_alimente_apprentissage(self):
        sig = dict(SIG, price=100.0, atr_m5=1.0)   # SL 98.5 / TP 104
        g = m.ghost_open(self.wb, sig, "régime BEAR")
        later = g["checked_ts"] + 60000
        with mock.patch.object(m, "fetch_candles",
                               lambda coin, iv, n, **k: [_m1(later, 100, 104.2, 99.8, 104)]):
            closed = m.ghost_update(self.wb)
        self.assertEqual(len(closed), 1)
        self.assertEqual(m._load_ghosts(), [])
        ws = self.wb["Fantômes"]
        self.assertEqual(ws.cell(2, 12).value, "TP")
        self.assertAlmostEqual(ws.cell(2, 14).value, 4.0 - m.GHOST_FEE_PCT)
        b = m._load_lessons()["buckets"]["regime_RANGING|buy|MR"]
        self.assertEqual(b["gw"], 1)
        self.assertNotIn("gl", b)

    def test_persistance_redemarrage(self):
        m.ghost_open(self.wb, dict(SIG, price=100.0, atr_m5=1.0), "x")
        m._ghosts = None   # simule un redémarrage
        self.assertEqual(len(m._load_ghosts()), 1)


def _declin(n, last_vol=300.0):
    """Bougies en baisse régulière (chaque bougie rouge), volume fort sur la dernière."""
    cs = []
    for i in range(n):
        px = 120.0 - 0.1 * i
        cs.append({"ts": i, "o": px + 0.05, "h": px + 0.06, "l": px - 0.01, "c": px, "v": 100.0})
    cs[-1]["v"] = last_vol
    return cs

class TestGhostCandidates(BaseTest):
    def scan(self, regime):
        m._ghost_candidates.clear(); m._reject_stats.clear()
        m._lessons_cache.clear(); m._lessons_cache_ts = 0.0
        with mock.patch.object(m, "fetch_candles", lambda coin, iv, n, **k: _declin(n)), \
             mock.patch.object(m, "_current_regime", regime):
            return m.get_signal("ETH")

    def test_mom_bloque_en_ranging_devient_candidat_fantome(self):
        self.assertIsNone(self.scan("RANGING"))
        self.assertEqual(len(m._ghost_candidates), 1)
        c = m._ghost_candidates[0]
        self.assertEqual((c["signal"], c["conditions"]["mode"]), ("sell", "MOM"))
        self.assertEqual(c["ghost_reason"], "MOM bloqué RANGING")
        self.assertEqual(m._reject_stats.get("MOM bloqué RANGING"), 1)

    def test_mom_en_bear_est_un_vrai_signal(self):
        sig = self.scan("BEAR")
        self.assertEqual((sig["signal"], sig["conditions"]["mode"]), ("sell", "MOM"))
        self.assertEqual(m._ghost_candidates, [])


def _mr_sell(iv, n, last_vol5):
    """SELL MR : M15 baissier, M5 en haut de bande avec 2 bougies rouges, M1 neutre."""
    cs = []
    if iv == m.INTERVAL_M15:
        for i in range(n):
            px = 130.0 - 0.2 * i
            cs.append({"ts": i, "o": px + 0.05, "h": px + 0.1, "l": px - 0.1, "c": px, "v": 100.0})
    elif iv == m.INTERVAL_M5:
        for i in range(n - 2):
            px = 100.0 + 0.1 * i
            cs.append({"ts": i, "o": px - 0.05, "h": px + 0.02, "l": px - 0.07, "c": px, "v": 100.0})
        top = cs[-1]["c"]
        cs.append({"ts": n - 2, "o": top + 0.02, "h": top + 0.03, "l": top - 0.03, "c": top - 0.02, "v": 100.0})
        cs.append({"ts": n - 1, "o": top - 0.02, "h": top - 0.01, "l": top - 0.06, "c": top - 0.05,
                   "v": last_vol5})
    else:
        for i in range(n):
            px = 100.0 + (0.05 if i % 2 else -0.05)
            cs.append({"ts": i, "o": 100.0, "h": px + 0.01, "l": px - 0.01, "c": px, "v": 100.0})
    return cs

class TestGhostVolume(BaseTest):
    def scan(self, last_vol5):
        m._ghost_candidates.clear(); m._reject_stats.clear()
        m._lessons_cache.clear(); m._lessons_cache_ts = 0.0
        with mock.patch.object(m, "fetch_candles", lambda coin, iv, n, **k: _mr_sell(iv, n, last_vol5)), \
             mock.patch.object(m, "_current_regime", "RANGING"), \
             mock.patch.object(m, "VOL_MULTIPLIER", 1.2):
            return m.get_signal("ETH")

    def test_mr_rejete_de_justesse_par_volume_devient_fantome_mrv(self):
        self.assertIsNone(self.scan(100.0))   # ratio 1.0 < 1.2
        self.assertEqual(len(m._ghost_candidates), 1)
        c = m._ghost_candidates[0]
        self.assertEqual((c["signal"], c["conditions"]["mode"]), ("sell", "MRv"))
        self.assertEqual(c["ghost_reason"], "MR volume ×1.00<1.2")
        self.assertEqual(m._reject_stats, {"volume M5": 1})

    def test_volume_trop_faible_pas_de_fantome(self):
        self.assertIsNone(self.scan(70.0))    # ratio < GHOST_VOL_MIN
        self.assertEqual(m._ghost_candidates, [])

    def test_volume_suffisant_vrai_signal_mr(self):
        sig = self.scan(300.0)
        self.assertEqual((sig["signal"], sig["conditions"]["mode"]), ("sell", "MR"))
        self.assertEqual(m._ghost_candidates, [])

    def test_fantome_mrv_clos_sans_toucher_apprentissage(self):
        if os.path.exists(m.GHOST_FILE):
            os.remove(m.GHOST_FILE)
        m._ghosts = None
        wb = m.init_excel()
        self.scan(100.0)
        g = m.ghost_open(wb, m._ghost_candidates[0], m._ghost_candidates[0]["ghost_reason"])
        later = g["checked_ts"] + 60000
        with mock.patch.object(m, "fetch_candles",
                               lambda coin, iv, n, **k: [_m1(later, g["entry"], g["sl"] + 1, g["entry"] - 0.01, g["entry"])]):
            closed = m.ghost_update(wb)
        self.assertEqual([x["status"] for x in closed], ["SL"])
        self.assertEqual(m._load_lessons().get("buckets", {}), {})
        self.assertEqual(wb["Fantômes"].cell(2, 5).value, "MRv")

    def test_heartbeat_par_type(self):
        closed = [{"side": "sell", "mode": "MOM", "pnl_pct": 1.0},
                  {"side": "sell", "mode": "MOM", "pnl_pct": -0.5},
                  {"side": "buy", "mode": "MRv", "pnl_pct": -0.3}]
        self.assertEqual(m._ghost_by_type(closed),
                         "BUY MRv ✓0/1 moy -0.30% · SELL MOM ✓1/2 moy +0.25%")


# ── Boucle principale (smoke test) ───────────────────────────────────────────
class _StopLoop(Exception):
    pass

class _Minute0(datetime):
    """Horloge figée à hh:00 : le résumé des leçons (toutes les 10 min) tombe à chaque cycle."""
    @classmethod
    def now(cls, tz=None):
        return datetime.now(tz).replace(minute=0)

def _at_utc_hour(h: int):
    """datetime dont now() tombe à h heures (même jour)."""
    class _H(datetime):
        @classmethod
        def now(cls, tz=None):
            return datetime.now(tz).replace(hour=h)
    return _H

class LoopInfo(FakeInfo):
    def meta(self): return {"universe": []}

class TestRunLoop(BaseTest):
    def run_loop(self, daily_chg, asset_positions=None, get_signal=None, d1_failures=0,
                 keep_regime=False, strategy="scalp", trend_closes=None):
        if os.path.exists(m.GHOST_FILE):
            os.remove(m.GHOST_FILE)
        m._ghosts = None
        m._lessons_cache.clear(); m._lessons_cache_ts = 0.0
        for attr in ("_sl_check_ts", "_prev_equity", "_prev_ts", "_ledger_ts") + \
                    (() if keep_regime else ("_regime_cache", "_regime_ts")):
            m.run.__dict__.pop(attr, None)
        info = LoopInfo(
            mids={c: "100" for c in m.WATCHLIST + m.TREND_COINS},
            perp={"assetPositions": asset_positions or [],
                  "crossMarginSummary": {"accountValue": "0", "totalMarginUsed": "0"}},
            spot={"balances": [{"coin": "USDC", "total": "111.18", "hold": "0"}]})
        ex = FakeExchange()
        cycles = []
        def fake_sleep(sec):
            if sec >= 5:                       # attente de fin de cycle
                cycles.append(sec)
                if len(cycles) >= 2:
                    raise _StopLoop()
        d1_calls, self.m1_calls = [], []
        def fc(coin, iv, n, **k):
            if iv != m.INTERVAL_D1:
                self.m1_calls.append(coin)
                return _declin(n)
            if n == max(m.TREND_LOOKBACKS) + 2:  # signal de la stratégie tendance
                return _d1(trend_closes or [100.0] * 60)
            d1_calls.append(coin)
            if len(d1_calls) <= d1_failures:
                raise RuntimeError("Bybit kline BTCUSDT: Too many visits")
            return _daily(daily_chg)
        with mock.patch.object(m, "info", info), mock.patch.object(m, "exchange", ex), \
             mock.patch.object(m, "fetch_candles", fc), \
             mock.patch.object(m.time, "sleep", fake_sleep), \
             mock.patch.object(m, "HEARTBEAT_HOURS", 1e-9), \
             mock.patch.object(m, "STRATEGY", strategy), \
             mock.patch.object(m, "get_signal", get_signal or m.get_signal):
            with self.assertRaises(_StopLoop):
                m.run()
        return ex, m._load_ghosts()

    def _signal_eth(self):
        """get_signal factice : BUY MR score 0.9 sur ETH (passe tous les filtres en RANGING)."""
        def gs(coin):
            if coin != "ETH":
                return None
            return {"symbol": "ETH", "signal": "buy", "score": 0.9, "atr_m5": 5.0, "price": 100.0,
                    "conditions": {"mode": "MR", "regime": "RANGING", "rsi": 50, "bbp": 0.3,
                                   "vol_ratio": 1.5}}
        return gs

    def test_signal_retenu_ouvre_un_ordre_en_mode_normal(self):
        """Témoin du test suivant : sans GHOST_ONLY le même signal passe un ordre (SL + TP)."""
        with mock.patch.object(m, "GHOST_ONLY", False):
            ex, _ = self.run_loop(-2, get_signal=self._signal_eth())
        self.assertTrue(ex.orders)

    def test_mode_fantome_seul_aucun_ordre(self):
        with mock.patch.object(m, "GHOST_ONLY", True):
            ex, ghosts = self.run_loop(-2, get_signal=self._signal_eth())
        self.assertEqual(ex.orders, [])
        self.assertEqual([(g["symbol"], g["reason"]) for g in ghosts],
                         [("ETH", "fantôme seul (aurait été tradé)")])
        self.assertIn("fantôme seul", dict(self.alerts)["🚀 Bot démarré"])
        self.assertIn("Positions scalping reprises : 0", dict(self.alerts)["🚀 Bot démarré"])
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertTrue(hb and hb[0].startswith("Mode : **fantôme seul**"))

    def test_wallet_partage_position_copytrade_ni_adoptee_ni_rescannee(self):
        """SOL ouvert par BotCopyTrade (absent de bot_state.json) : pas repris au démarrage
        (aucun SL posé, pas de fermeture zombie) et SOL exclu du scan réel."""
        scanned = []
        def gs(coin):
            scanned.append(coin)
            return None
        sol = {"position": {"coin": "SOL", "szi": "-0.5", "entryPx": "100", "unrealizedPnl": "-3"}}
        ex, _ = self.run_loop(-2, asset_positions=[sol], get_signal=gs)
        self.assertEqual(ex.orders, [])
        self.assertIn("⚠️ Position non reprise", self.titles())
        self.assertTrue(scanned)
        self.assertNotIn("SOL", scanned)
        self.assertIn("ETH", scanned)

    def test_bear_scan_fantome_et_heartbeat(self):
        """2 cycles complets de run() en BEAR : aucun ordre réel, SELL MOM suivi en fantôme,
        heartbeat Discord avec la ligne Fantômes."""
        ex, ghosts = self.run_loop(-10)
        self.assertEqual(ex.orders, [])                     # aucun ordre réel en BEAR
        self.assertTrue(ghosts)
        self.assertTrue(all(g["reason"] == "régime BEAR" for g in ghosts))
        self.assertTrue(any(g["mode"] == "MOM" and g["side"] == "sell" for g in ghosts))
        self.assertEqual(len(ghosts), len({(g["symbol"], g["side"], g["mode"]) for g in ghosts}))  # dédoublonné
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertTrue(hb and "Fantômes :" in hb[0])

    def test_regime_indisponible_reessaye_au_cycle_suivant(self):
        """Bougies D1 en échec au 1er cycle : le régime est recalculé dès le cycle suivant
        (avant : RANGING figé jusqu'à la prochaine demi-heure)."""
        ex, _ = self.run_loop(-10, d1_failures=1)
        self.assertEqual(ex.orders, [])
        hb = [d for t, d in self.alerts if "résumé" in t]      # 1 heartbeat par cycle
        self.assertIn("Régime : **RANGING**", hb[0])          # 1er cycle : défaut prudent
        self.assertIn("Régime : **BEAR**", hb[-1])            # 2e cycle : recalculé

    def test_regime_precedent_conserve_si_bybit_echoue(self):
        m.run._regime_cache = ("BEAR", "down", -10.0, -10.0)
        m.run._regime_ts = 0.0     # recalcul dû…
        ex, ghosts = self.run_loop(-2, d1_failures=99, keep_regime=True)   # …mais Bybit échoue
        self.assertTrue(ghosts and all(g["reason"] == "régime BEAR" for g in ghosts))

    def test_strategie_tendance_fantome_sans_scalping(self):
        """Mode tendance + GHOST_ONLY : une position fantôme par coin, aucun ordre, aucun scan
        M1/M5 du scalping, et le heartbeat résume les positions tendance."""
        with mock.patch.object(m, "GHOST_ONLY", True):
            ex, ghosts = self.run_loop(-2, strategy="trend", trend_closes=TestTrend.UP)
        self.assertEqual(ex.orders, [])
        self.assertEqual(self.m1_calls, [])
        self.assertEqual(ghosts, [])
        self.assertEqual(set(m._load_trend()["positions"]), set(m.TREND_COINS))
        self.assertIn("tendance journalière", dict(self.alerts)["🚀 Bot démarré"])
        self.assertIn("Coins : BTC, ETH, SOL, XRP, DOGE", dict(self.alerts)["🚀 Bot démarré"])
        self.assertNotIn("Positions scalping", dict(self.alerts)["🚀 Bot démarré"])
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertTrue(hb and "Tendance : BTC LONG +0.0 %" in hb[0])
        self.assertIn("Équité : **$111.18** · +0.0 % depuis le départ", dict(self.alerts)["🚀 Bot démarré"])
        self.assertIn("Latent : **+0.00 %** du capital", hb[0])
        self.assertIn("Équité : **$111.18** · +0.0 % depuis le départ", hb[0])

    def test_resume_des_lecons_du_scalping_absent_en_mode_tendance(self):
        """Le bloc « Leçons actives » (patterns du scalping, ~3 300 lignes de journal par jour) n'est
        plus affiché en mode tendance ; le mode scalping le garde."""
        shown = []
        with mock.patch.object(m, "datetime", _Minute0), \
             mock.patch.object(m, "print_lessons_summary", lambda: shown.append(m.STRATEGY)):
            with mock.patch.object(m, "GHOST_ONLY", True):
                self.run_loop(-2, strategy="trend", trend_closes=TestTrend.UP)
            self.assertEqual(shown, [])
            self.run_loop(-2)
        self.assertTrue(shown)
        self.assertEqual(set(shown), {"scalp"})

    def run_with_equity(self, equities, ledger_delta=None):
        """2 cycles en mode tendance fantôme ; get_equity renvoie `equities` (démarrage, cycle 1, cycle 2) et le
        journal Hyperliquid la ligne `ledger_delta`, datée du relevé. Retourne l'équité passée à trend_cycle."""
        eqs, seen = iter(equities), []
        ledger = lambda self_, addr, start: [{"time": start, "delta": ledger_delta}] if ledger_delta else []
        with mock.patch.object(m, "get_equity", lambda *a: next(eqs)), \
             mock.patch.object(m, "trend_cycle", lambda wb, us, eq: seen.append(eq) or []), \
             mock.patch.object(LoopInfo, "user_non_funding_ledger_updates", ledger), \
             mock.patch.object(m, "GHOST_ONLY", True):
            self.run_loop(-2, strategy="trend")
        return seen

    def test_depot_utilise_des_le_cycle_suivant_et_pas_compte_comme_gain(self):
        """Dépôt de 11.12 $ (+10 %) : tailles calculées sur la nouvelle équité dès le cycle suivant (avant :
        capital figé au démarrage tant que l'écart restait sous 20 %) ; le dépôt, lu dans le journal
        Hyperliquid, n'apparaît pas comme un gain dans le « % depuis le départ »."""
        send = {"type": "send", "user": "0x6b9e773128f453f5c2c60935ee2de2cbc5390a24",
                "destination": m.WALLET_ADDRESS.lower(), "token": "USDC", "amount": "11.12", "usdcValue": "11.12",
                "fee": "0.0"}
        seen = self.run_with_equity([111.18, 111.18, 122.30], send)
        self.assertEqual(seen, [111.18, 122.30])
        self.assertIn("💰 Dépôt de $11.12 reçu", self.titles())
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertIn("Équité : **$122.30** · +0.0 % depuis le départ", hb[-1])

    def test_retrait_pas_signale_comme_krach(self):
        """Retrait de 50 $ (−45 % en un cycle) : expliqué par le journal Hyperliquid, pas d'alerte « Flash
        Crash » (avant : alerte) ; une vraie chute sans retrait reste signalée."""
        withdraw = {"type": "withdraw", "usdc": "50.0", "nonce": 1, "fee": "1.0"}
        self.run_with_equity([111.18, 111.18, 61.18], withdraw)
        self.assertNotIn("⚡ Flash Crash Détecté", self.titles())
        self.assertIn("💸 Retrait de $50.00", self.titles())
        self.alerts.clear()
        self.run_with_equity([111.18, 111.18, 61.18])
        self.assertIn("⚡ Flash Crash Détecté", self.titles())

    def test_capital_fictif_du_test_annonce_et_utilise(self):
        """Mode fantôme avec PAPER_EQUITY=200 : tailles calculées sur 200 $ et capital fictif annoncé au démarrage
        et dans chaque résumé."""
        seen = []
        with mock.patch.object(m, "PAPER_EQUITY", 200.0), mock.patch.object(m, "GHOST_ONLY", True), \
             mock.patch.object(m, "trend_cycle", lambda wb, us, eq: seen.append(eq) or []):
            self.run_loop(-2, strategy="trend")
        self.assertEqual(seen, [200.0, 200.0])
        self.assertIn("capital fictif du test : **$200**", dict(self.alerts)["🚀 Bot démarré"])
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertTrue(hb[0].startswith("Mode : **fantôme seul** (aucun ordre réel) · capital fictif **$200**\n"))

    def test_erreur_imprevue_le_bot_continue(self):
        """Erreur imprévue au milieu d'un cycle : journalisée, une seule alerte, et les cycles suivants ont
        lieu (avant : bot arrêté ; systemd ne le relançait pas, la sortie semblant normale)."""
        with mock.patch.object(m, "calc_portfolio_pnl", mock.Mock(side_effect=KeyError("P&L"))), \
             mock.patch("sys.stderr", new_callable=io.StringIO):
            self.run_loop(-2)                                  # 2 cycles complets malgré l'erreur
        self.assertEqual(self.titles().count("⚠️ Erreur inattendue (le bot continue)"), 1)

    def test_crash_fatal_code_d_erreur_pour_systemd(self):
        """Erreur hors du cycle (démarrage) : alerte et code de sortie 1, pour que systemd (Restart=on-failure)
        relance le bot (avant : sortie 0, bot arrêté sans alerte). Ctrl+C : sortie normale."""
        with mock.patch.object(m, "run", mock.Mock(side_effect=RuntimeError("disque plein"))), \
             mock.patch("sys.stderr", new_callable=io.StringIO):
            with self.assertRaises(SystemExit) as cm:
                m.main()
        self.assertEqual(cm.exception.code, 1)
        self.assertIn("💥 Bot arrêté par une erreur", self.titles())
        with mock.patch.object(m, "run", mock.Mock(side_effect=KeyboardInterrupt)):
            m.main()

    def test_ranging_mom_bloque_suivi_en_fantome(self):
        ex, ghosts = self.run_loop(-2)
        self.assertEqual(ex.orders, [])
        self.assertTrue(ghosts)
        self.assertTrue(all(g["reason"] == "MOM bloqué RANGING" for g in ghosts))


# ── Régime de marché ─────────────────────────────────────────────────────────
def _daily(chg_pct, chg3_pct=None):
    """8 bougies D1 : variation totale chg_pct, et variation 3j chg3_pct."""
    start, end = 100.0, 100.0 * (1 + chg_pct / 100)
    closes = [start + (end - start) * i / 7 for i in range(8)]
    if chg3_pct is not None:
        closes[4] = end / (1 + chg3_pct / 100)
    return [{"c": c} for c in closes]

class TestRegime(BaseTest):
    def regime(self, btc, eth):
        data = {"BTC": btc, "ETH": eth}
        with mock.patch.object(m, "fetch_candles", lambda coin, iv, n, **k: data[coin]):
            return m.get_market_regime()[0]

    def test_bear_sur_la_moyenne(self):
        # Cas réel du 2026-10-14 : BTC -2.8 %, ETH -7.4 % → moyenne -5.1 % → BEAR
        self.assertEqual(self.regime(_daily(-2.8), _daily(-7.4)), "BEAR")

    def test_ranging(self):
        self.assertEqual(self.regime(_daily(-2.0), _daily(-5.9)), "RANGING")

    def test_bull(self):
        self.assertEqual(self.regime(_daily(6), _daily(7)), "BULL")

    def test_bull_en_pullback_3j_devient_ranging(self):
        self.assertEqual(self.regime(_daily(6, -3), _daily(7, -3)), "RANGING")

    def test_diverging(self):
        self.assertEqual(self.regime(_daily(4), _daily(-4)), "DIVERGING")

    def test_erreur_api_ranging_par_defaut(self):
        def boom(*a, **k): raise RuntimeError("api down")
        with mock.patch.object(m, "fetch_candles", boom):
            self.assertEqual(m.get_market_regime()[0], "RANGING")


# ── Signaux : diagnostic des rejets ──────────────────────────────────────────
def _candles(n, vol_last, vol=100.0, px=100.0):
    cs = [{"ts": i, "o": px, "h": px * 1.001, "l": px * 0.999, "c": px, "v": vol} for i in range(n)]
    cs[-1]["v"] = vol_last
    return cs

class TestSignalDiagnostic(BaseTest):
    def test_volume_bas_compte_comme_rejet_volume_m5(self):
        m._reject_stats.clear()
        def fc(coin, iv, n, **k): return _candles(n, vol_last=10.0)
        with mock.patch.object(m, "fetch_candles", fc):
            self.assertIsNone(m.get_signal("ETH"))
        self.assertEqual(m._reject_stats.get("volume M5"), 1)


# ── Dépôts et retraits (journal Hyperliquid) ────────────────────────────────
class TestDeposits(BaseTest):
    ME, OTHER = m.WALLET_ADDRESS.lower(), "0x6b9e773128f453f5c2c60935ee2de2cbc5390a24"

    def sync(self, ledger, state=None):
        m._state_atomic_write(state or {"initial_balance": 111.55, "ledger_ts": 1000})
        info = FakeInfo()
        info.ledger = ledger
        with mock.patch.object(m, "info", info):
            moves = m.sync_deposits()
        return moves, m._read_state()

    def send(self, t, amount, src=None, dst=None, token="USDC", fee="0.0"):
        return {"time": t, "delta": {"type": "send", "user": src or self.OTHER, "destination": dst or self.ME,
                                     "token": token, "amount": amount, "usdcValue": amount, "fee": fee}}

    def test_premier_releve_ne_recompte_pas_les_anciens_depots(self):
        """Premier relevé : les dépôts passés sont déjà dans la référence (111.55 $), seul le point de départ
        du journal est noté."""
        moves, state = self.sync([self.send(5000, "99.8")], {"initial_balance": 111.55})
        self.assertEqual(moves, [])
        self.assertEqual(state["initial_balance"], 111.55)
        self.assertGreater(state["ledger_ts"], 1_700_000_000_000)

    def test_depot_ajoute_a_la_reference_une_seule_fois(self):
        ledger = [self.send(2000, "90"),
                  self.send(3000, "5", src=self.ME, dst=self.ME),           # spot → perp du compte : interne
                  self.send(4000, "7", token="HYPE"),                       # autre jeton : hors équité suivie
                  {"time": 5000, "delta": {"type": "accountClassTransfer", "usdc": "3", "toPerp": True}}]
        moves, state = self.sync(ledger)
        self.assertEqual(moves, [(2000, 90.0)])
        self.assertEqual((state["initial_balance"], state["ledger_ts"]), (201.55, 5000))
        self.assertIn("$111.55 → **$201.55** (un dépôt n'est pas un gain)", dict(self.alerts)["💰 Dépôt de $90.00 reçu"])
        info = FakeInfo()
        info.ledger = ledger
        with mock.patch.object(m, "info", info):
            self.assertEqual(m.sync_deposits(), [])                        # déjà compté
        self.assertEqual(m._read_state()["initial_balance"], 201.55)

    def test_retrait_et_virement_sortant_retires_de_la_reference(self):
        ledger = [{"time": 2000, "delta": {"type": "withdraw", "usdc": "50.0", "nonce": 1, "fee": "1.0"}},
                  self.send(3000, "10", src=self.ME, dst=self.OTHER, fee="1.0"),
                  {"time": 4000, "delta": {"type": "deposit", "usdc": "20"}}]
        moves, state = self.sync(ledger)
        self.assertEqual([f for _, f in moves], [-50.0, -11.0, 20.0])
        self.assertAlmostEqual(state["initial_balance"], 111.55 - 41)
        self.assertEqual(self.titles(), ["💸 Retrait de $50.00", "💸 Retrait de $11.00", "💰 Dépôt de $20.00 reçu"])

    def test_journal_indisponible_rien_ne_change(self):
        m._state_atomic_write({"initial_balance": 111.55, "ledger_ts": 1000})
        info = FakeInfo()
        info.user_non_funding_ledger_updates = mock.Mock(side_effect=RuntimeError("502"))
        with mock.patch.object(m, "info", info):
            self.assertEqual(m.sync_deposits(), [])
        self.assertEqual(m._read_state(), {"initial_balance": 111.55, "ledger_ts": 1000})

    def test_operation_inconnue_signalee_dans_le_journal(self):
        moves, state = self.sync([{"time": 2000, "delta": {"type": "nouveauType", "usdc": "5"}}])
        self.assertEqual((moves, state["initial_balance"], state["ledger_ts"]), ([], 111.55, 2000))
        self.assertIn("« nouveauType » non comptée", sys.stdout.getvalue())

    def test_capital_fictif_du_test_ignore_en_reel(self):
        with mock.patch.object(m, "PAPER_EQUITY", 200.0):
            with mock.patch.object(m, "GHOST_ONLY", True):
                self.assertEqual(m._sizing_equity(111.19), 200.0)
            with mock.patch.object(m, "GHOST_ONLY", False):
                self.assertEqual(m._sizing_equity(111.19), 111.19)   # jamais pour un ordre réel

    def test_taille_sur_tout_le_solde_ou_le_plafond_du_env(self):
        self.assertEqual(m._sizing_equity(250.0), 250.0)                     # MAX_LIQUIDITY=0 : tout le solde
        with mock.patch.object(m, "_MAX_LIQ_CFG", 150.0):
            self.assertEqual(m._sizing_equity(250.0), 150.0)
            self.assertEqual(m._sizing_equity(90.0), 90.0)


# ── Utilitaires ──────────────────────────────────────────────────────────────
class TestUtils(BaseTest):
    def test_delai_max_des_requetes_hyperliquid(self):
        """Connexion restée bloquée (coupure Wi-Fi) : erreur au bout de HL_TIMEOUT_S au lieu d'un bot figé
        sans alerte (le SDK n'a aucun délai par défaut)."""
        self.assertEqual(m.info.kwargs.get("timeout"), m.HL_TIMEOUT_S)
        self.assertEqual(m.exchange.kwargs.get("timeout"), m.HL_TIMEOUT_S)
        self.assertLessEqual(m.HL_TIMEOUT_S, 30)

    def test_round_px(self):
        self.assertEqual(m.round_px(2012.345), 2012.3)
        self.assertEqual(m.round_px(0.08201357), 0.082014)

    def test_round_px_regle_hyperliquid(self):
        self.assertEqual(m.round_px(0.123456), 0.12346)    # DOGE ≥ 0.1 : 5 chiffres significatifs
        self.assertEqual(m.round_px(0.512345), 0.51235)    # XRP < 1
        self.assertEqual(m.round_px(9.876543), 9.8765)     # LTC/XRP < 10
        self.assertEqual(m.round_px(123456.7), 123457)     # entier toujours valide
        with mock.patch.dict(m._sym_filters, {"X": {"sz_dec": 4, "sz_step": 1e-4}}):
            self.assertEqual(m.round_px(1.23456, "X"), 1.23)   # au plus 6 − 4 = 2 décimales
        self.assertEqual(m.round_px(0.080481, "INCONNU"), 0.080481)   # filtres absents : pas de troncature

    def test_bybit_rate_limit_deux_echecs_puis_succes(self):
        """Bybit répond « Too many visits » quand les 15 requêtes du scan partent ensemble :
        2 nouvelles tentatives (avant : 1 seule → signal perdu pour le cycle)."""
        lim = {"retCode": 10006, "retMsg": "Too many visits. Exceeded the API Rate Limit."}
        ok  = {"retCode": 0, "result": {"list": [[str(i * 60_000), "1", "2", "0.5", "1.5", "10"]
                                                  for i in (3, 2, 1, 0)]}}
        calls = []
        def get(*a, **k):
            calls.append(1)
            r = mock.Mock(); r.json.return_value = lim if len(calls) < 3 else ok
            return r
        with mock.patch.object(m.requests, "get", get):
            cs = m.fetch_candles("ETH", m.INTERVAL_M1, 2)
        self.assertEqual(len(calls), 3)
        self.assertEqual([c["ts"] for c in cs], [60_000, 120_000])   # bougie en cours retirée

    def test_dashboard_lecons_comptent_les_fantomes(self):
        """Pattern appris uniquement en fantôme (1 gagné / 9 perdus → n=5, WR 10 %) : le bot
        le pénalise de 0.15 ; le dashboard affichait « 0 (apprentissage…) »."""
        m._lessons_cache = {"buckets": {"rsi70-75|buy|BBO": {"w": 0, "l": 0, "pnl": 0.0,
                                                             "gw": 1.0, "gl": 9.0}}}
        m._lessons_cache_ts = m.time.time()
        try:
            self.assertEqual(m.get_penalty("ETH", "buy", 72, 1.1, "RANGING", "BBO"), 0.15)
            m.write_dashboard(m.init_excel(), [], 100.0, 100.0, 0.0, 0.0, 0.0, 100.0,
                              "RANGING", 0.0, 0.0)
            with open("dashboard.html", encoding="utf-8") as f:
                html = f.read()
        finally:
            m._lessons_cache.clear(); m._lessons_cache_ts = 0.0
        row = html[html.index(">rsi70-75|buy|BBO<"):].split("</tr>")[0]
        self.assertIn("−0.15", row)
        self.assertIn("10%", row)

    def test_size_mult(self):
        self.assertEqual(m._size_mult(1.6), 1.15)
        self.assertEqual(m._size_mult(1.0), 1.0)
        self.assertEqual(m._size_mult(0.5), 0.85)

    def test_rsi_bornes(self):
        self.assertEqual(m._rsi([float(i) for i in range(30)], 14), 100.0)
        self.assertLess(m._rsi([float(30 - i) for i in range(30)], 14), 1.0)

    def test_bb_plat(self):
        self.assertEqual(m._bb([100.0] * 25)[3], 0.5)


class TestDiscordJournal(unittest.TestCase):
    """Chaque message Discord est copié dans EVENTS_FILE (relu par tools/status.py), même sans webhook."""
    def setUp(self):
        if os.path.exists(m.EVENTS_FILE):
            os.remove(m.EVENTS_FILE)

    def test_message_journalise_sans_webhook(self):
        with mock.patch.object(m, "DISCORD_WEBHOOK_URL", ""):
            m.send_discord_alert("💓 Bot actif", "ligne 1\nligne 2")
        with open(m.EVENTS_FILE, encoding="utf-8") as f:
            ev = [json.loads(l) for l in f]
        self.assertEqual(len(ev), 1)
        self.assertEqual((ev[0]["title"], ev[0]["description"]), ("💓 Bot actif", "ligne 1\nligne 2"))

    def test_limite_de_debit_discord_attente_puis_nouvel_essai(self):
        """Rafale de messages (minuit, passage en réel) : 429 → attente demandée par Discord puis nouvel
        envoi (avant : message perdu sans bruit) ; dernier envoi permis → attente avant le suivant."""
        resp = [mock.Mock(status_code=429, headers={"Retry-After": "1.5"}),
                mock.Mock(status_code=204, headers={"X-RateLimit-Remaining": "0", "X-RateLimit-Reset-After": "0.8"})]
        sent, slept = [], []
        with mock.patch.object(m, "DISCORD_WEBHOOK_URL", "http://hook"), \
             mock.patch.object(m.requests, "post", lambda *a, **k: sent.append(k) or resp.pop(0)), \
             mock.patch.object(m.time, "sleep", slept.append):
            m.send_discord_alert("t", "d")
        self.assertEqual(len(sent), 2)
        self.assertEqual(slept, [1.5, 0.8])

    def test_journal_en_erreur_n_empeche_pas_l_envoi(self):
        sent = []
        with mock.patch.object(m, "EVENTS_FILE", os.path.join(_TMP, "absent", "x.jsonl")), \
             mock.patch.object(m, "DISCORD_WEBHOOK_URL", "http://hook"), \
             mock.patch.object(m.requests, "post", lambda *a, **k: sent.append(k) or mock.Mock()), \
             mock.patch("sys.stdout", new_callable=io.StringIO):
            m.send_discord_alert("t", "d")
        self.assertEqual(len(sent), 1)



# ── Stratégie tendance journalière ───────────────────────────────────────────
def _d1(closes):
    """Bougies D1 closes se terminant hier (UTC), une par clôture (range ±2 % → ATR ≈ 4 % du prix)."""
    y, n = (int(m.time.time()) // 86400 - 1) * 86_400_000, len(closes)
    return [{"ts": y - (n - 1 - i) * 86_400_000, "o": c, "h": c * 1.02, "l": c * 0.98, "c": c, "v": 1.0}
            for i, c in enumerate(closes)]

class TrendExchange:
    """Hyperliquid simulé : journalise ouvertures, SL, fermetures et annulations."""
    def __init__(self, avg_px=100.0, sl=None, close_ok=True):
        self.avg_px, self.close_ok, self.sl, self.calls = avg_px, close_ok, sl or [_ok(21)], []
        self.stops = []   # (coin, qté, prix limite, déclenchement) de chaque ordre stop
    def update_leverage(self, *a, **k): return {"status": "ok"}
    def market_open(self, coin, is_buy, qty, px=None, slippage=0.02):
        self.calls.append(("open", coin, is_buy, qty))
        return {"status": "ok", "response": {"data": {"statuses": [
            {"filled": {"oid": 7, "avgPx": str(self.avg_px), "totalSz": str(qty)}}]}}}
    def order(self, coin, is_buy, qty, limit_px, order_type, reduce_only):
        trig = order_type["trigger"]
        self.calls.append((trig["tpsl"], coin, is_buy, trig["triggerPx"]))
        self.stops.append((coin, qty, limit_px, trig["triggerPx"]))
        return self.sl.pop(0) if len(self.sl) > 1 else self.sl[0]
    def market_close(self, coin, slippage=0.05):
        self.calls.append(("close", coin))
        return {"status": "ok"} if self.close_ok else {"status": "err"}
    def cancel(self, coin, oid):
        self.calls.append(("cancel", coin, oid))
        return {"status": "ok"}

def _held(coin, szi):
    return {"position": {"coin": coin, "szi": str(szi), "entryPx": "90", "unrealizedPnl": "0"}}

class TestTrend(BaseTest):
    UP   = [90.0] * 57 + [100.0]     # clôture > celles d'il y a 14, 28 et 56 jours → LONG
    DOWN = [110.0] * 57 + [100.0]    # → SHORT

    def setUp(self):
        super().setUp()
        self.wb = m.init_excel()
        m._trend_sl_ts = 1e18        # surveillance des SL testée à part
        self.fetches = []

    def cycle(self, ex=None, ghost=False, closes=None, per_coin=None, held=(), fills=None,
              mids=None, orders=None, equity=111.0, funding=(), rates=(), fetch=None, meta=None):
        perp = {"assetPositions": list(held)}
        info = OrdersInfo(orders=orders or [], mids=mids or {c: "100" for c in m.TREND_COINS},
                          perp=perp, fills=fills or [])
        info.funding, info.rates = list(funding), list(rates)
        if meta is not None:
            info.meta = meta
        ex = ex or TrendExchange()
        def fc(coin, iv, n, **k):
            self.fetches.append((coin, iv, n))
            return _d1((per_coin or {}).get(coin, closes or self.UP))
        with mock.patch.object(m, "info", info), mock.patch.object(m, "exchange", ex), \
             mock.patch.object(m, "fetch_candles", fetch or fc), mock.patch.object(m, "GHOST_ONLY", ghost):
            closed = m.trend_cycle(self.wb, perp, equity)
        return ex, closed

    def real_eth(self, side="BUY", **k):
        pos = {"id": "7", "symbol": "ETH", "side": side, "entry": 90.0, "qty": 0.12,
               "sl": 80.0 if side == "BUY" else 100.0, "sl_oid": 5, "atr": 3.0, "open_fee": 0.005,
               "opened_at": time.time() - 3 * 86400, "paper": False, **k}
        m._load_trend()["positions"]["ETH"] = pos
        return pos

    def eth_calls(self, ex):
        return [c for c in ex.calls if c[1] == "ETH"]

    # ── Signal ──
    def test_signal_long_short_une_requete_par_coin_et_par_jour(self):
        with mock.patch.object(m, "fetch_candles",
                               lambda c, iv, n, **k: self.fetches.append((c, iv, n)) or
                               _d1(self.UP if c == "ETH" else self.DOWN)):
            self.assertEqual(m.trend_signal("ETH")[0], "BUY")
            self.assertEqual(m.trend_signal("SOL")[0], "SELL")
            side, close, atr = m.trend_signal("ETH")
        self.assertEqual(self.fetches, [("ETH", "1D", 58), ("SOL", "1D", 58)])
        self.assertEqual(close, 100.0)
        self.assertGreater(atr, 2.0)

    def test_signal_attend_la_bougie_d_hier(self):
        """Juste après minuit, Bybit n'a pas encore publié la bougie d'hier : pas de signal
        (sinon décision sur l'avant-veille), nouvel essai au cycle suivant."""
        stale = [dict(c, ts=c["ts"] - 86_400_000) for c in _d1(self.UP)]
        with mock.patch.object(m, "fetch_candles", lambda *a, **k: self.fetches.append(a) or stale):
            self.assertIsNone(m.trend_signal("ETH"))
            self.assertIsNone(m.trend_signal("ETH"))
        self.assertEqual(len(self.fetches), 2)

    def test_signal_vote_majoritaire_14_28_56_jours(self):
        """LONG si la clôture dépasse celle d'il y a N jours pour au moins 2 horizons sur 3 :
        hausse sur 14 et 56 j, baisse sur 28 j → LONG (la règle 28 j seule disait SHORT)."""
        closes = [100.0] * 58
        closes[-15], closes[-29], closes[-57] = 90.0, 110.0, 90.0
        with mock.patch.object(m, "fetch_candles", lambda *a, **k: _d1(closes)):
            self.assertEqual(m.trend_signal("ETH")[0], "BUY")
            closes[-57] = 110.0                                  # baisse sur 28 et 56 j → SHORT
            m._trend_sig.clear()
            self.assertEqual(m.trend_signal("ETH")[0], "SELL")

    def test_quantite_minimum_hyperliquid_et_arrondi(self):
        with mock.patch.dict(m._sym_filters, {"XRP": {"sz_dec": 0, "sz_step": 1}}):
            self.assertEqual(m._trend_qty("XRP", 2.5, 111.0), 5)     # 11.10 $ → 4.44 → arrondi au-dessus
            self.assertEqual(m._trend_qty("XRP", 2.5, 1000.0), 40)   # 100 $ pile
            self.assertEqual(m._trend_qty("XRP", 2.5, 50.0), 0)      # 10,5 $ (+1 %) > 10 $ par coin → refus
        # Pire prix d'un ordre à 88 (ex. limite du SL d'un long) : 10,5 $ + 1 % à ce prix → 0.121 au lieu de 0.111
        self.assertEqual(m._trend_qty("ETH", 100.0, 111.0, 88.0), 0.121)
        self.assertEqual(m._trend_qty("ETH", 100.0, 111.0, 112.0), 0.111)   # au-dessus du prix : sans effet

    # ── Mode fantôme ──
    def test_fantome_ouvre_une_position_par_coin_sans_ordre(self):
        ex, _ = self.cycle(ghost=True, per_coin={"SOL": self.DOWN})
        self.assertEqual(ex.calls, [])
        pos = m._load_trend()["positions"]
        self.assertEqual({c: p["side"] for c, p in pos.items()},
                         {"BTC": "BUY", "ETH": "BUY", "DOGE": "BUY", "XRP": "BUY", "SOL": "SELL"})
        self.assertTrue(all(p["paper"] for p in pos.values()))
        self.assertLess(pos["ETH"]["sl"], 100.0)
        self.assertGreater(pos["SOL"]["sl"], 100.0)
        self.assertAlmostEqual(abs(pos["ETH"]["sl"] - 100.0), m.TREND_SL_ATR * pos["ETH"]["atr"], delta=0.01)
        rows = list(self.wb["Tendance"].iter_rows(min_row=2, values_only=True))
        self.assertEqual([r[4] for r in rows], ["Fantôme"] * 5)
        self.cycle(ghost=True, per_coin={"SOL": self.DOWN})          # même signal : rien de nouveau
        self.assertEqual(self.wb["Tendance"].max_row, 6)
        m._trend = None                                               # redémarrage
        self.assertEqual(set(m._load_trend()["positions"]), set(m.TREND_COINS))

    def test_fantome_sl_touche_puis_a_plat_jusqu_a_demain(self):
        self.cycle(ghost=True)
        sl = m._load_trend()["positions"]["ETH"]["sl"]
        mids = {c: "100" for c in m.TREND_COINS}
        mids["ETH"] = str(sl - 0.5)
        _, closed = self.cycle(ghost=True, mids=mids)
        self.assertEqual([(p["symbol"], p["status"]) for p in closed], [("ETH", "SL touché")])
        self.assertLess(closed[0]["pnl"], 0)
        _, closed = self.cycle(ghost=True)
        self.assertEqual(closed, [])
        self.assertNotIn("ETH", m._load_trend()["positions"])

    # ── Réel ──
    def test_inversion_du_signal_ferme_puis_ouvre_en_sens_inverse(self):
        self.real_eth()
        fills = [{"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "1.2", "fee": "0.005"}]
        ex, closed = self.cycle(closes=self.DOWN, held=[_held("ETH", 0.12)], fills=fills)
        calls = self.eth_calls(ex)
        self.assertEqual([c[0] for c in calls], ["close", "cancel", "open", "sl"])
        self.assertEqual(calls[1], ("cancel", "ETH", 5))
        self.assertFalse(calls[2][2])                    # vente
        self.assertTrue(calls[3][2] and calls[3][3] > 100.0)   # SL d'un short : rachat au-dessus
        self.assertNotIn("tp", [c[0] for c in ex.calls])       # pas de TP : sortie à l'inversion
        self.assertEqual([(p["symbol"], p["status"]) for p in closed], [("ETH", "signal inversé")])
        self.assertAlmostEqual(closed[0]["pnl"], 1.2 - 0.005 - 0.005)
        new = m._load_trend()["positions"]["ETH"]
        self.assertEqual((new["side"], new["sl_oid"], new["paper"]), ("SELL", 21, False))
        self.assertAlmostEqual(new["open_fee"], m.TREND_FEE * new["qty"] * 100)   # absents de la réponse : estimés

    def test_sl_reel_touche_pas_de_reouverture_le_meme_jour(self):
        """SL exécuté en 2 fills : P&L et prix de sortie moyen enregistrés (avant : sortie à 0)."""
        self.real_eth()
        fills = [{"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "-1.0", "fee": "0.003", "px": "80.2", "sz": "0.08"},
                 {"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "-0.5", "fee": "0.002", "px": "79.9", "sz": "0.04"}]
        ex, closed = self.cycle(fills=fills)              # ETH absent du compte : SL exécuté
        self.assertEqual([(p["symbol"], p["status"]) for p in closed], [("ETH", "SL touché")])
        self.assertAlmostEqual(closed[0]["pnl"], -1.51)
        self.assertAlmostEqual(closed[0]["exit"], 80.1)
        self.assertIn("Entrée : $90 → sortie $80.1", dict(self.alerts)["🛑 ETH BUY — SL touché [tendance]"])
        self.assertEqual(self.eth_calls(ex), [("cancel", "ETH", 5)])   # sans effet : SL déjà exécuté
        ex, _ = self.cycle()
        self.assertEqual(self.eth_calls(ex), [])
        self.assertEqual(len([c for c in ex.calls if c[0] == "open"]), 0)   # les 4 autres déjà ouverts

    def test_fill_de_fermeture_absent_attend_3_cycles(self):
        """Toujours aucun fill après 3 cycles : P&L estimé au prix du SL (avant : 0 $, « Fermé (SL) »)."""
        pos = self.real_eth()
        for _ in range(2):
            _, closed = self.cycle()
            self.assertEqual(closed, [])
            self.assertIn("ETH", m._load_trend()["positions"])
        _, closed = self.cycle()
        self.assertEqual([(p["symbol"], p["status"], p["exit"]) for p in closed], [("ETH", "SL touché", 80.0)])
        self.assertAlmostEqual(closed[0]["pnl"], 0.12 * (80 - 90) - m.TREND_FEE * 0.12 * 80 - 0.005, places=4)

    def test_sl_d_un_jour_passe_reouverture_sans_attendre(self):
        """SL exécuté hier pendant un arrêt du bot : la journée d'attente est déjà passée, la position
        est rouverte dès ce cycle (avant : à plat jusqu'à demain)."""
        self.real_eth()
        fills = [{"coin": "ETH", "dir": "Close Long", "time": (time.time() - 86400) * 1000, "oid": 9,
                  "closedPnl": "-1.5", "fee": "0.005", "px": "80", "sz": "0.12"}]
        ex, closed = self.cycle(fills=fills)
        self.assertEqual([p["status"] for p in closed], ["SL touché"])
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["cancel", "open", "sl"])
        self.assertNotIn("ETH", m._load_trend()["stopped"])

    def test_prix_absent_nouvel_essai_au_cycle_suivant(self):
        """Prix ETH absent de Hyperliquid un instant : rien ce cycle, ouverture au suivant (avant :
        « équité insuffisante » et plus d'essai avant demain)."""
        mids = {c: "100" for c in m.TREND_COINS if c != "ETH"}
        ex, _ = self.cycle(mids=mids)
        self.assertEqual(self.eth_calls(ex), [])
        self.assertNotIn("⚠️ Tendance : ETH non ouvert", self.titles())
        ex, _ = self.cycle()
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["open", "sl"])

    def test_position_enregistree_des_son_ouverture(self):
        """Bot arrêté au milieu des ouvertures : celles déjà faites sont dans trend_state.json (avant :
        écrit en fin de cycle seulement → position réelle prise ensuite pour celle d'un autre bot)."""
        ex = TrendExchange()
        def open_(coin, *a, **k):
            if coin == "ETH":
                raise KeyboardInterrupt
            return TrendExchange.market_open(ex, coin, *a, **k)
        ex.market_open = open_
        with self.assertRaises(KeyboardInterrupt):
            self.cycle(ex=ex)
        with open(m.TREND_FILE, encoding="utf-8") as f:
            self.assertEqual(list(json.load(f)["positions"]), ["BTC"])

    def test_fermeture_sans_fill_visible_pnl_estime(self):
        """Fermeture faite mais fills pas encore visibles : P&L estimé au prix moyen, frais d'ouverture
        compris (avant : 0 $)."""
        self.real_eth(open_fee=m.TREND_FEE * 0.12 * 90)
        _, closed = self.cycle(closes=self.DOWN, held=[_held("ETH", 0.12)])
        self.assertEqual([(p["status"], p["exit"]) for p in closed], [("signal inversé", 100.0)])
        self.assertAlmostEqual(closed[0]["pnl"], 1.2 - m.TREND_FEE * 0.12 * (100 + 90), places=4)

    def test_coin_tenu_par_un_autre_bot_jamais_ouvert(self):
        ex, _ = self.cycle(held=[_held("SOL", -0.5)])
        self.assertEqual(sorted(c[1] for c in ex.calls if c[0] == "open"), ["BTC", "DOGE", "ETH", "XRP"])
        self.assertNotIn("SOL", m._load_trend()["positions"])

    def test_passage_en_reel_n_ouvre_pas_un_coin_tenu_par_un_autre_bot(self):
        """Fin du mode fantôme : SOL fantôme est fermé, mais SOL est tenu sur le compte par un autre
        bot (BotCopyTrade) → pas de SOL réel, qui fusionnerait avec sa position (avant : ouvert)."""
        self.cycle(ghost=True)
        ex, closed = self.cycle(held=[_held("SOL", -0.5)])
        self.assertIn(("SOL", "fin du mode fantôme"), [(p["symbol"], p["status"]) for p in closed])
        self.assertEqual(sorted(c[1] for c in ex.calls if c[0] == "open"), ["BTC", "DOGE", "ETH", "XRP"])
        self.assertNotIn("SOL", m._load_trend()["positions"])

    # ── Capital fictif du mode fantôme (test) ──
    def test_capital_fictif_du_test_positions_fantomes_recalculees(self):
        """Capital fictif de 200 $ : les positions fantômes en cours sont recalculées une seule fois à cette
        taille (même entrée, même SL, ligne Excel à jour), une position réelle n'est pas touchée ; au passage
        en réel, le % du capital de leurs fermetures se calcule sur ces 200 $."""
        self.cycle(ghost=True, per_coin={"SOL": self.DOWN})                    # 111 $ : environ 11 $ par coin
        before = {c: dict(p) for c, p in m._load_trend()["positions"].items()}
        self.real_eth()
        held = [_held("ETH", 0.12)]
        with mock.patch.object(m, "PAPER_EQUITY", 200.0):
            for _ in range(2):
                self.cycle(ghost=True, equity=200.0, per_coin={"SOL": self.DOWN}, held=held)
        pos = m._load_trend()["positions"]
        rows = {r[0]: r for r in self.wb["Tendance"].iter_rows(min_row=2, values_only=True)}
        for c in ("BTC", "SOL", "XRP", "DOGE"):
            self.assertEqual((pos[c]["entry"], pos[c]["sl"]), (before[c]["entry"], before[c]["sl"]))
            self.assertAlmostEqual(pos[c]["qty"] * pos[c]["entry"], 20.0, delta=0.11)   # 50 % × 200 $ / 5 coins
            self.assertAlmostEqual(pos[c]["open_fee"], m.TREND_FEE * pos[c]["qty"] * pos[c]["entry"])
            self.assertEqual(rows[pos[c]["id"]][5], pos[c]["qty"])
        self.assertEqual(pos["ETH"]["qty"], 0.12)                                  # réelle : intacte
        self.assertEqual(self.titles().count("🧪 Test : capital fictif de $200"), 1)
        self.assertIn("BTC, DOGE, SOL, XRP", dict(self.alerts)["🧪 Test : capital fictif de $200"])
        _, closed = self.cycle(equity=111.0, per_coin={"SOL": self.DOWN}, held=held)   # passage en réel
        sol = [p for p in closed if p["symbol"] == "SOL"][0]
        self.assertEqual(sol["status"], "fin du mode fantôme")
        self.assertIn(f"P&L : **{sol['pnl'] / 200 * 100:+.2f} %** du capital",
                      dict(self.alerts)["👻 🛑 SOL SELL — fin du mode fantôme [tendance]"])

    # ── Réponses perdues (délai dépassé, 5xx, bot arrêté pendant un ordre) ──
    def lost_open(self, coin="ETH"):
        """Hyperliquid dont la réponse à l'ouverture de `coin` se perd (exception après l'envoi)."""
        ex = TrendExchange()
        def open_(c, is_buy, qty, **k):
            if c != coin:
                return TrendExchange.market_open(ex, c, is_buy, qty, **k)
            ex.calls.append(("open", c, is_buy, qty))
            raise requests.exceptions.ReadTimeout("Read timed out. (read timeout=20)")
        ex.market_open = open_
        return ex

    def held_now(self, skip=()):
        """Le compte tel que le bot le suit (positions réelles), sauf les coins `skip`."""
        return [_held(c, p["qty"] if p["side"] == "BUY" else -p["qty"])
                for c, p in m._load_trend()["positions"].items() if c not in skip]

    def test_ouverture_sans_reponse_position_reprise_avec_son_sl(self):
        """Ordre d'ouverture passé mais réponse perdue : au cycle suivant, même après un redémarrage, la
        position trouvée sur le compte est reprise et reçoit son SL, sans second ordre (avant : prise pour
        celle d'un autre bot, jamais gérée, sans SL)."""
        ex, _ = self.cycle(ex=self.lost_open())
        (_, _, _, qty), = self.eth_calls(ex)                       # ordre parti, pas de SL
        self.assertNotIn("ETH", m._load_trend()["positions"])
        self.assertIn("ETH", m.trend_coins())                       # ni adopté par le scalping
        self.assertIn("n'a pas répondu", dict(self.alerts)["⚠️ Tendance : ETH non ouvert"])
        self.assertNotIn("ETH", m._load_trend()["stopped"])
        m._trend = None                                             # redémarrage du bot
        ex, _ = self.cycle(held=self.held_now() + [_held("ETH", qty)])
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["sl"])
        pos = m._load_trend()["positions"]["ETH"]
        self.assertEqual((pos["side"], pos["qty"], pos["entry"], pos["sl_oid"], pos["paper"]), ("BUY", qty, 90.0, 21, False))
        self.assertEqual(pos["sl"], m.round_px(90.0 - m.TREND_SL_ATR * pos["atr"], "ETH"))
        self.assertEqual(m._load_trend()["opening"], {})
        self.assertIn("position retrouvée sur le compte", dict(self.alerts)["🟢 LONG ETH ouvert [tendance]"])
        ex, _ = self.cycle(held=self.held_now())
        self.assertEqual(self.eth_calls(ex), [])

    def test_ouverture_sans_reponse_non_passee_nouvel_essai_au_cycle_suivant(self):
        """Réponse perdue et rien sur le compte (l'ordre n'est pas arrivé) : nouvel essai dès le cycle suivant
        (avant : plus d'essai avant demain)."""
        self.cycle(ex=self.lost_open())
        ex, _ = self.cycle(held=self.held_now())
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["open", "sl"])
        self.assertEqual(m._load_trend()["positions"]["ETH"]["sl_oid"], 21)
        self.assertEqual(m._load_trend()["opening"], {})

    def test_renfort_sans_reponse_la_taille_du_compte_fait_foi(self):
        """Renfort passé mais réponse perdue (fermeture pas faite) : au cycle suivant le compte montre le
        renfort, la fermeture part sans second renfort et ses frais sont estimés (avant : second renfort)."""
        self.real_eth(side="SELL")                                   # 0.12 ETH vendus à 90
        mids = {c: "100" for c in m.TREND_COINS}
        mids["ETH"] = "80"
        fills = [{"coin": "ETH", "dir": "Close Short", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "2.56", "fee": "0.009"}]
        ex, closed = self.cycle(held=[_held("ETH", -0.256)], fills=fills, mids=mids)   # signal LONG : inversion
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["close", "cancel", "open", "sl"])
        self.assertAlmostEqual(closed[0]["top_up"], 0.136)
        self.assertAlmostEqual(closed[0]["pnl"], 2.56 - 0.005 - 0.009 - m.TREND_FEE * 0.136 * 80, places=4)

    def test_fermeture_sans_reponse_sl_orphelin_annule(self):
        """Fermeture passée mais réponse perdue : la position reste suivie ; au cycle suivant, même après un
        redémarrage, absente du compte, sa fermeture est enregistrée avec son vrai motif, son SL resté en place
        est annulé (avant : SL orphelin, qui aurait fermé une future position du même sens à l'ancien niveau)
        et la position inverse s'ouvre aussitôt (avant : « Fermé (SL) », à plat jusqu'au lendemain)."""
        self.real_eth()
        ex = TrendExchange()
        def close_(coin, slippage=0.05):
            ex.calls.append(("close", coin))
            raise requests.exceptions.ReadTimeout("Read timed out.")
        ex.market_close = close_
        _, closed = self.cycle(ex=ex, closes=self.DOWN, held=[_held("ETH", 0.12)])
        self.assertEqual(closed, [])
        self.assertIn("n'a pas répondu", dict(self.alerts)["⚠️ Tendance : fermeture ETH sans réponse"])
        m._trend = None                                             # redémarrage du bot
        fills = [{"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "1.2", "fee": "0.005", "px": "100", "sz": "0.12"}]
        ex, closed = self.cycle(closes=self.DOWN, fills=fills, held=self.held_now(skip=("ETH",)))
        self.assertEqual([(p["symbol"], p["status"], round(p["pnl"], 2)) for p in closed],
                         [("ETH", "signal inversé", 1.19)])
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["cancel", "open", "sl"])
        self.assertEqual(self.eth_calls(ex)[0], ("cancel", "ETH", 5))
        self.assertEqual(m._load_trend()["positions"]["ETH"]["side"], "SELL")
        self.assertNotIn("ETH", m._load_trend()["stopped"])

    def test_fermeture_sans_reponse_non_passee_un_sl_ensuite_reste_un_sl(self):
        """Réponse perdue mais ordre jamais arrivé (position toujours sur le compte), puis signal revenu : un
        SL touché ensuite est enregistré comme un SL (à plat jusqu'à demain), pas comme cette fermeture."""
        self.real_eth()
        ex = TrendExchange()
        ex.market_close = mock.Mock(side_effect=requests.exceptions.ConnectionError("Wi-Fi coupé"))
        self.cycle(ex=ex, per_coin={"ETH": self.DOWN}, held=[_held("ETH", 0.12)])
        m._trend_sig.clear()                                         # le lendemain : signal revenu LONG
        ex, _ = self.cycle(held=self.held_now())
        self.assertEqual(self.eth_calls(ex), [])
        self.assertNotIn("closing", m._load_trend()["positions"]["ETH"])
        fills = [{"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 5,
                  "closedPnl": "-1.2", "fee": "0.005", "px": "80", "sz": "0.12"}]
        ex, closed = self.cycle(fills=fills, held=self.held_now(skip=("ETH",)))
        self.assertEqual([p["status"] for p in closed], ["SL touché"])
        self.assertEqual(self.eth_calls(ex), [("cancel", "ETH", 5)])
        self.assertIn("ETH", m._load_trend()["stopped"])

    def test_annulation_du_sl_en_echec_reessayee_au_cycle_suivant(self):
        """Position fermée mais annulation de son SL en échec (réseau) : nouvel essai au cycle suivant, même
        après un redémarrage (avant : SL orphelin laissé sur Hyperliquid)."""
        self.real_eth()
        ex = TrendExchange()
        ex.cancel = mock.Mock(side_effect=requests.exceptions.ConnectionError("Wi-Fi coupé"))
        fills = [{"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "-1.2", "fee": "0.005", "px": "100", "sz": "0.12"}]
        _, closed = self.cycle(ex=ex, closes=self.DOWN, held=[_held("ETH", 0.12)], fills=fills)
        self.assertEqual([p["status"] for p in closed], ["signal inversé"])
        self.assertEqual(m._load_trend()["cancels"], [["ETH", 5]])
        m._trend = None                                             # redémarrage du bot
        stop = {"coin": "ETH", "oid": 5, "isTrigger": True, "reduceOnly": True, "orderType": "Stop Market", "side": "A"}
        ex, _ = self.cycle(closes=self.DOWN, held=self.held_now(), orders=[stop])
        self.assertEqual(ex.calls, [("cancel", "ETH", 5)])
        self.assertEqual(m._load_trend()["cancels"], [])
        m._load_trend()["cancels"].append(["ETH", 5])               # SL déjà parti entre-temps : rien à refaire
        ex, _ = self.cycle(closes=self.DOWN, held=self.held_now())
        self.assertEqual((ex.calls, m._load_trend()["cancels"]), ([], []))

    # ── Aucun ordre sous TREND_MIN_NOTIONAL (10,5 $), à son pire prix ──
    def assert_ordres_au_minimum(self, ex, px=100.0):
        """Chaque ouverture (vente : jusqu'à 2 % sous px) et chaque SL (vente : jusqu'à sa limite ;
        achat : à son déclenchement) vaut au moins TREND_MIN_NOTIONAL."""
        for _, coin, is_buy, qty in [c for c in ex.calls if c[0] == "open"]:
            self.assertGreaterEqual(qty * (px if is_buy else px * (1 - m.TREND_SLIP_OPEN)), m.TREND_MIN_NOTIONAL, coin)
        for coin, qty, limit_px, trigger in ex.stops:
            self.assertGreaterEqual(qty * min(limit_px, trigger), m.TREND_MIN_NOTIONAL, coin)

    def test_long_ouverture_et_sl_au_minimum_a_leur_pire_prix(self):
        """Le SL d'un long (vente jusqu'à 10 % sous son déclenchement) vaut ≥ 10,5 $ à ce prix limite,
        au cas où Hyperliquid y contrôle son minimum de 10 $ (avant le 2026-11-04 : 0.126 × 78.66 = 9.91 $)."""
        ex, _ = self.cycle()
        self.assertEqual(len(ex.stops), 5)
        self.assert_ordres_au_minimum(ex)

    def test_short_ouverture_au_minimum_a_son_pire_prix(self):
        """Équité de 100 $ (10 $ visés par coin) : la vente d'ouverture d'un short vaut ≥ 10,5 $ (+1 % de
        marge) même exécutée 2 % sous le prix → 0.109 (avant : minimum de 11 $ au prix moyen → 0.11)."""
        ex, _ = self.cycle(closes=self.DOWN, equity=100.0)
        self.assertEqual([c[3] for c in ex.calls if c[0] == "open"], [0.109] * 5)
        self.assert_ordres_au_minimum(ex)

    def test_short_tres_gagnant_renforce_avant_fermeture(self):
        """Short dont la fermeture vaudrait 9.60 $ (0.12 × 80) : renfort d'abord (vente ≥ 10,5 $ à son pire
        prix), puis fermeture du tout ; frais du renfort déduits du P&L (avant : ordre de 9.60 $)."""
        self.real_eth(side="SELL")                        # 0.12 ETH vendus à 90
        mids = {c: "100" for c in m.TREND_COINS}
        mids["ETH"] = "80"
        fills = [{"coin": "ETH", "dir": "Close Short", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "1.2", "fee": "0.005"}]
        ex, closed = self.cycle(held=[_held("ETH", -0.12)], fills=fills, mids=mids)   # signal LONG : inversion
        calls = self.eth_calls(ex)
        self.assertEqual([c[0] for c in calls], ["open", "close", "cancel", "open", "sl"])
        add = calls[0][3]
        self.assertFalse(calls[0][2])                                   # renfort du short : une vente
        self.assertGreaterEqual(add * 80 * (1 - m.TREND_SLIP_OPEN), m.TREND_MIN_NOTIONAL)
        self.assertGreaterEqual((0.12 + add) * 80, m.TREND_MIN_NOTIONAL)   # fermeture du tout (achat)
        self.assertEqual([p["status"] for p in closed], ["signal inversé"])
        self.assertAlmostEqual(closed[0]["pnl"], 1.2 - 0.005 - 0.005 - m.TREND_FEE * add * 80, places=4)
        self.assertIn(f"Renforcée de {add:g}", dict(self.alerts)["✅ ETH SELL — signal inversé [tendance]"])
        self.assertEqual(m._load_trend()["positions"]["ETH"]["side"], "BUY")

    def test_renfort_au_minimum_malgre_l_arrondi_du_prix_limite(self):
        """DOGE à 0.002481 : le SDK arrondit la limite de vente 0.00243138 à 0.002431 (6 décimales). Un renfort
        visé pile à 10,5 $ (4319) y valait 10.4995 $ (trouvé par tests/test_trend_sim) ; avec 1 % de marge, ≥ 10,5 $."""
        pos = {"id": "7", "symbol": "DOGE", "side": "SELL", "entry": 0.003, "qty": 3700.0, "sl": 0.0034,
               "sl_oid": 5, "atr": 0.0001, "open_fee": 0.005, "opened_at": time.time(), "paper": False}
        ex = TrendExchange()
        with mock.patch.object(m, "exchange", ex), \
             mock.patch.dict(m._sym_filters, {"DOGE": {"sz_dec": 0, "sz_step": 1}}):
            m._trend_top_up(pos, 0.002481)
        add   = ex.calls[0][3]
        limit = round(float(f"{0.002481 * (1 - m.TREND_SLIP_OPEN):.5g}"), 6)   # calculée comme le SDK
        self.assertGreaterEqual(add * limit, m.TREND_MIN_NOTIONAL)

    def test_renfort_fait_mais_fermeture_refusee_pas_de_second_renfort(self):
        """Renfort passé, fermeture refusée : au cycle suivant la position (renfort compris) vaut assez,
        on retente la fermeture seule — jamais de renforts en série."""
        self.real_eth(side="SELL")
        mids = {c: "100" for c in m.TREND_COINS}
        mids["ETH"] = "80"
        ex = TrendExchange(close_ok=False)
        for szi in (-0.12, -0.256):                     # compte avant puis après le renfort
            self.cycle(ex=ex, held=[_held("ETH", szi)], mids=mids)
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["open", "close", "close"])
        self.assertAlmostEqual(m._load_trend()["positions"]["ETH"]["top_up"], 0.136)

    # ── Sources de bougies, tailles des coins ──
    def hl_candles(self, closes):
        """candles_snapshot de Hyperliquid renvoyant les bougies D1 `closes` (format de l'API)."""
        hl = [{"t": c["ts"], "o": str(c["o"]), "h": str(c["h"]), "l": str(c["l"]), "c": str(c["c"]), "v": "1"}
              for c in _d1(closes)]
        return lambda coin, iv, start, end: [c for c in hl if start <= c["t"] <= end]

    def test_bybit_indisponible_signal_avec_les_bougies_hyperliquid(self):
        """Bybit injoignable (panne, accès bloqué) ou figé : signal calculé sur les bougies journalières
        de Hyperliquid (avant : plus aucune décision, positions figées sans alerte)."""
        info = FakeInfo()
        info.candles_snapshot = self.hl_candles(self.DOWN)
        with mock.patch.object(m, "fetch_candles", mock.Mock(side_effect=RuntimeError("bybit"))), \
             mock.patch.object(m, "info", info):
            self.assertEqual(m.trend_signal("ETH")[0], "SELL")
            m._trend_sig.clear()
            info.candles_snapshot = self.hl_candles(self.UP)
            self.assertEqual(m.trend_signal("ETH")[0], "BUY")
        # Bybit répond mais sans la bougie d'hier (données figées) : Hyperliquid aussi
        stale = [dict(c, ts=c["ts"] - 86_400_000) for c in _d1(self.UP)]
        m._trend_sig.clear()
        info.candles_snapshot = self.hl_candles(self.DOWN)
        with mock.patch.object(m, "fetch_candles", lambda *a, **k: stale), mock.patch.object(m, "info", info):
            self.assertEqual(m.trend_signal("ETH")[0], "SELL")

    def test_aucune_bougie_alerte_une_fois_passe_1_h(self):
        """Ni Bybit ni Hyperliquid : rien n'est fait (positions gardées avec leur SL) ; une alerte par coin
        et par jour passé 1 h UTC (juste après minuit, la bougie peut simplement tarder)."""
        down = mock.Mock(side_effect=RuntimeError("bybit"))
        with mock.patch.object(m, "datetime", _at_utc_hour(0)):
            ex, _ = self.cycle(fetch=down)
        self.assertEqual((ex.calls, self.titles()), ([], []))
        with mock.patch.object(m, "datetime", _at_utc_hour(3)):
            self.cycle(fetch=down)
            self.cycle(fetch=down)
        self.assertEqual(self.titles().count("⚠️ Tendance : pas de signal ETH"), 1)
        self.assertEqual(len(self.titles()), 5)

    def test_tailles_inconnues_aucune_ouverture_reelle_puis_nouvel_essai(self):
        """Tailles des coins non chargées (Hyperliquid injoignable au démarrage) : aucune ouverture réelle
        (quantité au mauvais pas → refusée, ou BTC à 0.001 = 78 $), nouvel essai à chaque cycle, puis
        ouvertures dès qu'elles sont connues (avant : mauvaises tailles toute la session)."""
        universe = {"universe": [{"name": c, "szDecimals": 3} for c in m.TREND_COINS]}
        with mock.patch.object(m, "_sym_filters_ok", False), mock.patch.dict(m._sym_filters, {}, clear=True):
            ex, _ = self.cycle(meta=mock.Mock(side_effect=RuntimeError("502")))
            self.assertEqual([c for c in ex.calls if c[0] == "open"], [])
            self.assertFalse(m._sym_filters_ok)
            ex, _ = self.cycle(meta=lambda: universe)
            self.assertEqual(len([c for c in ex.calls if c[0] == "open"]), 5)
            self.assertTrue(m._sym_filters_ok)

    # ── Financement compris dans le P&L ──
    def test_financement_reel_compte_dans_le_pnl(self):
        """Position réelle fermée : P&L des fills + financement payé sur CE coin depuis l'ouverture (avant :
        ignoré, alors que le backtest le compte)."""
        pos = self.real_eth()
        t0 = pos["opened_at"] * 1000
        funding = [{"time": t0 - 3_600_000, "delta": {"coin": "ETH", "usdc": "-9"}},     # avant l'ouverture
                   {"time": t0 + 3_600_000, "delta": {"coin": "ETH", "usdc": "-0.03"}},
                   {"time": t0 + 7_200_000, "delta": {"coin": "BTC", "usdc": "-5"}},     # autre coin
                   {"time": t0 + 7_200_000, "delta": {"coin": "ETH", "usdc": "-0.02"}}]
        fills = [{"coin": "ETH", "dir": "Close Long", "time": time.time() * 1000, "oid": 9,
                  "closedPnl": "1.2", "fee": "0.005", "px": "100", "sz": "0.12"}]
        _, closed = self.cycle(closes=self.DOWN, held=[_held("ETH", 0.12)], fills=fills, funding=funding)
        self.assertAlmostEqual(closed[0]["pnl"], 1.2 - 0.005 - 0.005 - 0.05)
        self.assertAlmostEqual(closed[0]["funding"], -0.05)
        self.assertIn("Financement payé : $0.05 (compris dans le P&L)",
                      dict(self.alerts)["✅ ETH BUY — signal inversé [tendance]"])

    def test_financement_fantome_estime_avec_les_taux_publics(self):
        """Short fantôme : reçoit le financement quand le taux est positif (les longs paient les shorts)."""
        self.cycle(ghost=True, closes=self.DOWN)
        pos = dict(m._load_trend()["positions"]["ETH"])
        rates = [{"time": pos["opened_at"] * 1000 + h * 3_600_000, "fundingRate": "0.0000125"} for h in range(1, 25)]
        m._trend_sig.clear()                                             # nouveau jour : signal LONG, inversion
        _, closed = self.cycle(ghost=True, rates=rates)
        eth = [p for p in closed if p["symbol"] == "ETH"][0]
        funding = pos["qty"] * pos["entry"] * 24 * 0.0000125
        self.assertAlmostEqual(eth["funding"], funding, places=4)
        self.assertAlmostEqual(eth["pnl"], m._paper_pnl(pos, 100.0) + funding, places=4)

    def test_financement_sur_plusieurs_pages_ou_indisponible(self):
        """Position tenue des semaines : plus de 500 paiements, toutes les pages sont lues. API en erreur :
        financement non compté (None), la fermeture n'est pas bloquée."""
        pos = self.real_eth()
        t0 = int(pos["opened_at"] * 1000)
        rows = [{"time": t0 + i, "delta": {"coin": "ETH", "usdc": "-0.001"}} for i in range(1, 701)]
        info, starts = FakeInfo(), []
        info.user_funding_history = lambda a, start: starts.append(start) or [r for r in rows if r["time"] >= start][:500]
        with mock.patch.object(m, "info", info):
            self.assertAlmostEqual(m._trend_funding(pos), -0.7)
            self.assertEqual(starts, [t0, t0 + 501])
            info.user_funding_history = mock.Mock(side_effect=RuntimeError("502"))
            self.assertIsNone(m._trend_funding(pos))

    # ── Liste des coins (BTC remplace LTC le 2026-11-03) ──
    MIDS6 = {c: "100" for c in ("BTC", "ETH", "SOL", "XRP", "DOGE", "LTC")}

    # ── Messages Discord en % du capital ──
    def test_ouverture_et_sl_en_pourcentages_du_capital(self):
        self.cycle(ghost=True, per_coin={"SOL": self.DOWN})
        pos = m._load_trend()["positions"]["SOL"]
        opened = dict(self.alerts)["👻 🔴 SHORT SOL ouvert [tendance]"]
        risk = pos["qty"] * (pos["sl"] - pos["entry"]) + m.TREND_FEE * pos["qty"] * (pos["entry"] + pos["sl"])
        self.assertIn("Notionnel : **$11.1** (10.0 % du capital)", opened)
        self.assertIn(f"perte au SL : **{-risk / 111 * 100:.2f} %** du capital", opened)
        mids = {c: "100" for c in m.TREND_COINS}
        mids["SOL"] = str(pos["sl"] + 1)
        _, closed = self.cycle(ghost=True, per_coin={"SOL": self.DOWN}, mids=mids)
        msg = dict(self.alerts)["👻 🛑 SOL SELL — SL touché [tendance]"]
        self.assertIn(f"P&L : **{closed[0]['pnl'] / 111 * 100:+.2f} %** du capital · ", msg)
        self.assertAlmostEqual(closed[0]["pnl"] / 111 * 100, -risk / 111 * 100, places=2)

    def test_bilan_heartbeat_en_pourcentages_du_capital(self):
        m._load_trend()["positions"]["ETH"] = {"symbol": "ETH", "side": "BUY", "qty": 0.111, "entry": 100.0,
                                               "sl": 90.0, "paper": True}
        ws = self.wb["Tendance"]
        ws.append(["1", "", "SOL/USDT", "SELL", "Fantôme", 0.1, 100, 110, 11.1, "SL touché", 110, -1.11, -10.0, ""])
        ws.append(["2", "", "BTC/USDT", "BUY", "Réel", 0.1, 100, 90, 11.1, "signal inversé", 150, 5.0, 45.0, ""])
        ws.append(["3", "", "XRP/USDT", "BUY", "Fantôme", 0.1, 100, 90, 11.1, "Ouvert", None, None, None, None])
        with mock.patch.object(m, "GHOST_ONLY", True):
            txt = m.trend_heartbeat(self.wb, {"ETH": "120"}, 111.0, [{"pnl": -1.11}])
        self.assertIn("Tendance : ETH LONG +20.0 %", txt)
        self.assertIn("Latent : **+2.00 %** du capital", txt)
        self.assertIn("Fermetures tendance : 1 · **-1.00 %** du capital ($-1.11)", txt)
        self.assertIn("Bilan fantôme depuis le début : **+1.00 %** du capital (1 fermée, 0 gagnante)", txt)

    def test_coin_retire_de_la_liste_ferme_puis_jamais_rouvert(self):
        self.cycle(ghost=True, mids=self.MIDS6)
        m._load_trend()["positions"]["LTC"] = dict(m._load_trend()["positions"]["ETH"], symbol="LTC")
        _, closed = self.cycle(ghost=True, mids=self.MIDS6)
        self.assertEqual([(p["symbol"], p["status"]) for p in closed], [("LTC", "retiré de la stratégie")])
        self.assertEqual(set(m._load_trend()["positions"]), {"BTC", "ETH", "SOL", "XRP", "DOGE"})
        _, closed = self.cycle(ghost=True, mids=self.MIDS6)
        self.assertEqual(closed, [])

    def test_coin_retire_en_reel_ferme_au_marche_et_sl_annule(self):
        m._load_trend()["positions"]["LTC"] = dict(self.real_eth(), symbol="LTC")
        del m._load_trend()["positions"]["ETH"]
        ex, closed = self.cycle(held=[_held("LTC", 0.12)], mids=self.MIDS6)
        self.assertEqual([c for c in ex.calls if c[1] == "LTC"], [("close", "LTC"), ("cancel", "LTC", 5)])
        self.assertEqual([p["status"] for p in closed], ["retiré de la stratégie"])
        self.assertNotIn("LTC", m._load_trend()["positions"])

    def test_tailles_chargees_pour_les_coins_tendance_et_scalping(self):
        meta = {"universe": [{"name": c, "szDecimals": d} for c, d in (("BTC", 5), ("LTC", 2), ("AVAX", 2))]}
        with mock.patch.object(m, "info", mock.Mock(meta=lambda: meta)), \
             mock.patch.dict(m._sym_filters, {}, clear=True):
            m.init_exchange_info()
            self.assertEqual((m._flt("BTC")["sz_dec"], m._flt("BTC")["sz_step"]), (5, 1e-05))
            self.assertIn("LTC", m._sym_filters)        # liste du scalping toujours servie
            self.assertNotIn("AVAX", m._sym_filters)

    def test_compte_illisible_aucune_ouverture_reelle(self):
        ex = TrendExchange()
        info = OrdersInfo(mids={c: "100" for c in m.TREND_COINS})
        with mock.patch.object(m, "info", info), mock.patch.object(m, "exchange", ex), \
             mock.patch.object(m, "fetch_candles", lambda *a, **k: _d1(self.UP)), \
             mock.patch.object(m, "GHOST_ONLY", False):
            m.trend_cycle(self.wb, None, 111.0)
        self.assertEqual(ex.calls, [])

    def test_fin_du_mode_fantome_bascule_en_reel(self):
        self.cycle(ghost=True)
        ex, closed = self.cycle()
        self.assertEqual({p["status"] for p in closed}, {"fin du mode fantôme"})
        self.assertEqual(len(closed), 5)
        self.assertEqual(len([c for c in ex.calls if c[0] == "open"]), 5)
        self.assertFalse(any(p["paper"] for p in m._load_trend()["positions"].values()))

    def test_position_reelle_gardee_si_retour_en_mode_fantome(self):
        self.real_eth()
        ex, _ = self.cycle(ghost=True, held=[_held("ETH", 0.12)])
        self.assertEqual(self.eth_calls(ex), [])
        self.assertFalse(m._load_trend()["positions"]["ETH"]["paper"])

    def test_sl_refuse_deux_fois_referme_et_attend_demain(self):
        ex = TrendExchange(sl=[_err()])
        ex, _ = self.cycle(ex=ex, closes=None, per_coin=None)
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["open", "sl", "sl", "close"])
        self.assertEqual(m._load_trend()["positions"], {})
        self.assertEqual(self.titles().count("⚠️ Tendance : ETH non ouvert"), 1)
        n = len(ex.calls)
        self.cycle(ex=ex)
        self.assertEqual(len(ex.calls), n)                           # pas de nouvel essai aujourd'hui
        self.assertEqual(self.titles().count("⚠️ Tendance : ETH non ouvert"), 1)

    def test_fermeture_refusee_garde_la_position_et_n_ouvre_pas_l_inverse(self):
        self.real_eth()
        ex, closed = self.cycle(ex=TrendExchange(close_ok=False), closes=self.DOWN, held=[_held("ETH", 0.12)])
        self.assertEqual([c[0] for c in self.eth_calls(ex)], ["close"])
        self.assertEqual(closed, [])
        self.assertEqual(m._load_trend()["positions"]["ETH"]["side"], "BUY")
        self.assertNotIn("closing", m._load_trend()["positions"]["ETH"])   # refus net : rien à vérifier ensuite
        self.assertIn("⚠️ Tendance : fermeture ETH échouée", self.titles())

    def test_sl_tendance_manquant_repose(self):
        self.real_eth()
        m._trend_sl_ts = 0.0
        ex, _ = self.cycle(held=[_held("ETH", 0.12)], orders=[])
        self.assertIn(("sl", "ETH", False, 80.0), ex.calls)
        self.assertEqual(m._load_trend()["positions"]["ETH"]["sl_oid"], 21)

    def test_reprise_scalping_ignore_les_positions_tendance(self):
        """Une position tendance n'est ni adoptée par le scalping ni signalée « non reprise », même si
        bot_state.json garde une ancienne position scalping du même sens sur ce coin (avant : adoptée)."""
        self.real_eth()
        m._save_trend()
        m._trend = None
        for saved in ([], [{"symbol": "ETH", "side": "BUY", "id": "old", "sl": 85.0, "tp": 99.0}]):
            m.save_positions_state(saved)
            with mock.patch.object(m, "info", FakeInfo(perp={"assetPositions": [_held("ETH", 0.12)]})):
                rec, _ = m.recover_open_positions()
            self.assertEqual(rec, [])
        self.assertEqual(self.titles(), [])


if __name__ == "__main__":
    unittest.main()
