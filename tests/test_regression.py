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
    "LEVERAGE": "5", "HEARTBEAT_HOURS": "6",
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


class BaseTest(unittest.TestCase):
    def setUp(self):
        for f in (m.EXCEL_FILE, m.STATE_FILE, m.TREND_FILE):
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


# ── Excel ────────────────────────────────────────────────────────────────────
class TestSecuriteEtJournal(BaseTest):
    def _load_main(self, env):
        """Second exemplaire de main.py, importé avec un autre environnement."""
        import importlib.util
        spec = importlib.util.spec_from_file_location("main_bis", os.path.join(ROOT, "main.py"))
        mod = importlib.util.module_from_spec(spec)
        with mock.patch.dict(os.environ, env):
            spec.loader.exec_module(mod)
        return mod

    def test_cle_du_compte_par_defaut(self):
        mod = self._load_main({"HL_ACCOUNT_ADDRESS": ""})
        self.assertEqual(mod.WALLET_ADDRESS, mod._account.address)
        self.assertEqual(mod.exchange.kwargs["account_address"], mod._account.address)

    def test_cle_deleguee_le_compte_suivi_est_celui_de_l_adresse(self):
        """Clé d'API wallet (trade sans pouvoir retirer) : positions, équité et ordres portent sur le compte
        HL_ACCOUNT_ADDRESS, pas sur l'adresse de la clé (avant : seule la clé du compte était acceptée)."""
        addr = "0x" + "ab" * 20
        mod = self._load_main({"HL_ACCOUNT_ADDRESS": f" {addr} "})
        self.assertEqual(mod.WALLET_ADDRESS, addr)
        self.assertNotEqual(mod._account.address.lower(), addr)
        self.assertEqual(mod.exchange.kwargs["account_address"], addr)

    def test_env_lisible_par_son_proprietaire_seulement(self):
        with open("env_test", "w") as f:
            f.write("X=1\n")
        os.chmod("env_test", 0o664)
        try:
            m._protect_env_file("env_test")
            self.assertEqual(os.stat("env_test").st_mode & 0o777, 0o600)
            m._protect_env_file("absent")                      # fichier absent : sans erreur
        finally:
            os.remove("env_test")

    def test_erreurs_copiees_dans_un_fichier_date(self):
        """Le journal systemd ne garde qu'environ 1,5 jour : chaque erreur est aussi écrite, datée, dans
        bot_errors.log ; fichier trop gros → renommé en .1."""
        for f in (m.ERROR_LOG, m.ERROR_LOG + ".1"):
            if os.path.exists(f):
                os.remove(f)
        try:
            m.log_error("user_state: 502")
            line = open(m.ERROR_LOG, encoding="utf-8").read()
            self.assertRegex(line, r"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ user_state: 502\n$")
            with mock.patch.object(m, "ERROR_LOG_BYTES", 10):
                m.log_error("suivante")
            self.assertEqual(open(m.ERROR_LOG + ".1", encoding="utf-8").read(), line)
            self.assertIn("suivante", open(m.ERROR_LOG, encoding="utf-8").read())
        finally:
            for f in (m.ERROR_LOG, m.ERROR_LOG + ".1"):
                if os.path.exists(f):
                    os.remove(f)


class TestExcel(BaseTest):
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

    def test_ancien_fichier_garde_ses_onglets(self):
        """Fichier existant avec les onglets de l'ancien scalping : historique gardé tel quel, onglet
        Tendance ajouté une seule fois."""
        wb = openpyxl.Workbook(); wb.active.title = "Trades"; wb.active.append(["ID", "x"])
        wb.save(m.EXCEL_FILE)
        self.assertEqual(m.init_excel().sheetnames, ["Trades", "Tendance"])
        wb2 = m.init_excel()
        self.assertEqual(wb2.sheetnames, ["Trades", "Tendance"])
        self.assertEqual(wb2["Trades"].cell(1, 2).value, "x")

    def test_nouveau_fichier_onglet_tendance_seul(self):
        self.assertEqual(m.init_excel().sheetnames, ["Tendance"])


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


# ── Fills de fermeture : P&L réel d'une position ─────────────────────────────
class TestFees(BaseTest):
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

    def test_closing_pnl_frais_ouverture_reels(self):
        gross, open_fee, close_fee = m._closing_pnl(self.DOGE_POS, self.DOGE_FILLS)
        self.assertAlmostEqual(gross, -0.295832, places=6)
        self.assertAlmostEqual(open_fee, 0.036116, places=6)
        self.assertAlmostEqual(close_fee, 0.035987, places=6)


class TestClosingAttribution(BaseTest):
    def test_fermeture_anterieure_a_l_ouverture_ignoree(self):
        fills = [{"coin": "DOGE", "dir": "Close Short", "time": 1_000_000, "oid": 9,
                  "closedPnl": "5.0", "fee": "0.02"}]
        pos = {"id": "1", "symbol": "DOGE", "opened_at": 2_000.0}   # ouvert après ce fill
        self.assertIsNone(m._closing_pnl(pos, fills))


# ── Surveillance des SL ──────────────────────────────────────────────────────
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


# ── Boucle principale ────────────────────────────────────────────────────────
class _StopLoop(Exception):
    pass

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
    def run_loop(self, trend_closes=None, cycles=2):
        """`cycles` cycles complets de run(), avec un résumé Discord à chaque cycle ; retourne l'Exchange simulé."""
        for attr in ("_prev_equity", "_prev_ts", "_ledger_ts"):
            m.run.__dict__.pop(attr, None)
        info = LoopInfo(
            mids={c: "100" for c in m.TREND_COINS},
            perp={"assetPositions": [], "crossMarginSummary": {"accountValue": "0", "totalMarginUsed": "0"}},
            spot={"balances": [{"coin": "USDC", "total": "111.18", "hold": "0"}]})
        ex, ends = TrendExchange(), []
        def fake_sleep(sec):
            if sec > m.STARTUP_RETRY_S:        # attente de fin de cycle (≈ 60 s)
                ends.append(sec)
                if len(ends) >= cycles:
                    raise _StopLoop()
        with mock.patch.object(m, "info", info), mock.patch.object(m, "exchange", ex), \
             mock.patch.object(m, "fetch_candles", lambda coin, iv, n, **k: _d1(trend_closes or [100.0] * 60)), \
             mock.patch.object(m.time, "sleep", fake_sleep), \
             mock.patch.object(m, "HEARTBEAT_HOURS", 1e-9):
            with self.assertRaises(_StopLoop):
                m.run()
        return ex

    def test_mode_fantome_une_position_par_coin_sans_ordre(self):
        """Mode fantôme : une position fantôme par coin et aucun ordre ; message de démarrage et résumé Discord
        de la stratégie tendance, sans la ligne « Régime » de l'ancien scalping."""
        with mock.patch.object(m, "GHOST_ONLY", True):
            ex = self.run_loop(trend_closes=TestTrend.UP)
        self.assertEqual(ex.calls, [])
        self.assertEqual(set(m._load_trend()["positions"]), set(m.TREND_COINS))
        start = dict(self.alerts)["🚀 Bot démarré"]
        self.assertIn("fantôme seul", start)
        self.assertIn("tendance journalière", start)
        self.assertIn("Coins : BTC, ETH, SOL, XRP, DOGE", start)
        self.assertIn("Équité : **$111.18** · +0.0 % depuis le départ", start)
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertTrue(hb and hb[0].startswith("Mode : **fantôme seul**"))
        self.assertIn("Tendance : BTC LONG +0.0 %", hb[0])
        self.assertIn("Latent : **+0.00 %** du capital", hb[0])
        self.assertIn("Équité : **$111.18** · +0.0 % depuis le départ", hb[0])
        self.assertNotIn("Régime", hb[0])

    def test_mode_reel_chaque_coin_ouvert_avec_son_sl(self):
        with mock.patch.object(m, "GHOST_ONLY", False):
            ex = self.run_loop(trend_closes=TestTrend.UP, cycles=1)
        self.assertEqual(sorted(c[1] for c in ex.calls if c[0] == "open"), sorted(m.TREND_COINS))
        self.assertEqual(sorted(c[1] for c in ex.calls if c[0] == "sl"), sorted(m.TREND_COINS))
        self.assertNotIn("fantôme", dict(self.alerts)["🚀 Bot démarré"])

    def run_with_equity(self, equities, ledger_delta=None):
        """2 cycles en mode fantôme ; get_equity renvoie (ou lève) `equities` dans l'ordre, du démarrage aux
        cycles, et le journal Hyperliquid la ligne `ledger_delta`, datée du relevé. Retourne l'équité passée
        à trend_cycle."""
        eqs, seen = iter(equities), []
        def get_equity(*a):
            v = next(eqs)
            if isinstance(v, Exception):
                raise v
            return v
        ledger = lambda self_, addr, start: [{"time": start, "delta": ledger_delta}] if ledger_delta else []
        with mock.patch.object(m, "get_equity", get_equity), \
             mock.patch.object(m, "trend_cycle", lambda wb, us, eq: seen.append(eq) or []), \
             mock.patch.object(LoopInfo, "user_non_funding_ledger_updates", ledger), \
             mock.patch.object(m, "GHOST_ONLY", True):
            self.run_loop()
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

    def test_equite_illisible_au_demarrage_aucune_taille_supposee(self):
        """Hyperliquid injoignable au démarrage puis au 1er cycle : le bot attend l'équité réelle avant toute
        décision (avant : 100 $ supposés, et les positions du 1er cycle taillées sur ce montant)."""
        seen = self.run_with_equity([RuntimeError("502"), RuntimeError("502"), 111.18, 111.18, 111.18])
        self.assertEqual(seen, [111.18, 111.18])
        self.assertEqual(sys.stdout.getvalue().count("équité au démarrage"), 2)

    def test_capital_fictif_du_test_annonce_et_utilise(self):
        """Mode fantôme avec PAPER_EQUITY=200 : tailles calculées sur 200 $ et capital fictif annoncé au démarrage
        et dans chaque résumé."""
        seen = []
        with mock.patch.object(m, "PAPER_EQUITY", 200.0), mock.patch.object(m, "GHOST_ONLY", True), \
             mock.patch.object(m, "trend_cycle", lambda wb, us, eq: seen.append(eq) or []):
            self.run_loop()
        self.assertEqual(seen, [200.0, 200.0])
        self.assertIn("capital fictif du test : **$200**", dict(self.alerts)["🚀 Bot démarré"])
        hb = [d for t, d in self.alerts if "résumé" in t]
        self.assertTrue(hb[0].startswith("Mode : **fantôme seul** (aucun ordre réel) · capital fictif **$200**\n"))

    def test_erreur_imprevue_le_bot_continue(self):
        """Erreur imprévue au milieu d'un cycle : journalisée, une seule alerte, et les cycles suivants ont
        lieu (avant : bot arrêté ; systemd ne le relançait pas, la sortie semblant normale)."""
        with mock.patch.object(m, "trend_heartbeat", mock.Mock(side_effect=KeyError("P&L"))), \
             mock.patch.object(m, "GHOST_ONLY", True), mock.patch("sys.stderr", new_callable=io.StringIO):
            self.run_loop()                                  # 2 cycles complets malgré l'erreur
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


# ── Tableau de bord ──────────────────────────────────────────────────────────
class TestDashboard(BaseTest):
    def test_positions_equite_et_latent(self):
        m._load_trend()["positions"]["BTC"] = {"id": "1", "symbol": "BTC", "side": "SELL", "entry": 76111.0,
                                               "qty": 0.00028, "sl": 82806.0, "opened_at": 1793904000.0,
                                               "paper": False}
        st = {"assetPositions": [{"position": {"coin": "BTC", "szi": "-0.00028", "unrealizedPnl": "0.8355"}}]}
        m.write_dashboard(m.init_excel(), st, 220.56, 211.3483)
        with open("dashboard.html", encoding="utf-8") as f:
            html = f.read()
        self.assertIn("<b>$220.56</b>", html)
        self.assertIn("+4.4 % depuis le départ", html)
        self.assertIn("<td>BTC</td><td>SHORT</td><td>0.00028</td><td>$76111</td><td>$82806</td>", html)
        self.assertIn("$+0.84 (+0.38 %)", html)

    def test_erreur_journalisee_sans_arreter_le_bot(self):
        m.write_dashboard(None, None, 220.0, 211.0)          # classeur absent : erreur interne
        self.assertIn("write_dashboard", sys.stdout.getvalue())


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
        """Bybit répond « Too many visits » (limite de débit) : 2 nouvelles tentatives (avant : 1 seule →
        signal perdu pour le cycle)."""
        lim = {"retCode": 10006, "retMsg": "Too many visits. Exceeded the API Rate Limit."}
        ok  = {"retCode": 0, "result": {"list": [[str(i * 60_000), "1", "2", "0.5", "1.5", "10"]
                                                  for i in (3, 2, 1, 0)]}}
        calls = []
        def get(*a, **k):
            calls.append(1)
            r = mock.Mock(); r.json.return_value = lim if len(calls) < 3 else ok
            return r
        with mock.patch.object(m.requests, "get", get):
            cs = m.fetch_candles("ETH", m.INTERVAL_D1, 2)
        self.assertEqual(len(calls), 3)
        self.assertEqual([c["ts"] for c in cs], [60_000, 120_000])   # bougie en cours retirée


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
        self.assertIn("ETH", m._load_trend()["opening"])            # noté : vérifié au cycle suivant
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

    def test_tailles_chargees_pour_tous_les_coins(self):
        """Pas de taille de chaque coin de Hyperliquid : une position sur un coin retiré de TREND_COINS peut
        encore être renforcée puis fermée au bon pas."""
        meta = {"universe": [{"name": c, "szDecimals": d} for c, d in (("BTC", 5), ("LTC", 2), ("AVAX", 2))]}
        with mock.patch.object(m, "info", mock.Mock(meta=lambda: meta)), \
             mock.patch.dict(m._sym_filters, {}, clear=True):
            self.assertTrue(m.init_exchange_info())
            self.assertEqual((m._flt("BTC")["sz_dec"], m._flt("BTC")["sz_step"]), (5, 1e-05))
            self.assertEqual(set(m._sym_filters), {"BTC", "LTC", "AVAX"})

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


if __name__ == "__main__":
    unittest.main()
