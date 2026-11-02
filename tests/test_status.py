"""
Tests de tools/status.py — fichiers locaux factices, sans réseau ni journal système réel.
"""
import json, os, sys, tempfile, unittest
from unittest import mock
from datetime import datetime, timezone

import test_regression as T          # isolation (stubs, réseau interdit) EN PREMIER
sys.path.insert(0, os.path.join(T.ROOT, "tools"))
import status as S                    # noqa: E402
from openpyxl import Workbook


class TestStatus(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="bst_status_")

    def test_stats(self):
        self.assertEqual(S.stats([1.0, -1.0, 2.0, 2.0]), "n=   4  WR= 75%  moy=+1.000%  t=+1.4")

    def test_read_events_filtre_par_date_et_ignore_lignes_cassees(self):
        p = os.path.join(self.dir, "ev.jsonl")
        with open(p, "w", encoding="utf-8") as f:
            f.write(json.dumps({"t": "2026-10-27T10:00:00+00:00", "title": "vieux"}) + "\n")
            f.write("{pas du json\n")
            f.write(json.dumps({"t": "2026-10-28T10:00:00+00:00", "title": "récent"}) + "\n")
        ev = S.read_events(p, datetime(2026, 10, 28, tzinfo=timezone.utc))
        self.assertEqual([e["title"] for e in ev], ["récent"])
        self.assertEqual(S.read_events(os.path.join(self.dir, "absent"), datetime.now(timezone.utc)), [])

    def test_bilan_fantomes_separe_aurait_ete_trade(self):
        wb = Workbook(); wb.active.title = "Trades"
        ws = wb.create_sheet("Fantômes")
        ws.append(["ID", "Ouvert (UTC)", "Symbole", "Côté", "Mode", "Régime", "Score", "Raison",
                   "Entrée", "SL", "TP", "Statut", "Sortie", "P&L %", "Fermé (UTC)"])
        row = lambda side, mode, reason, pnl, closed, st="TP": \
            ["G", "", "ETH/USDT", side, mode, "BULL", 1, reason, 1, 1, 1, st, 1, pnl, closed]
        ws.append(row("BUY", "BBO", "fantôme seul (aurait été tradé)", 0.5, "2026-10-28 08:00"))
        ws.append(row("BUY", "BBO", "BBO bloqué RANGING", -0.2, "2026-10-28 08:00"))
        ws.append(row("SELL", "MOM", "fantôme seul (aurait été tradé)", 0.3, "2026-10-01 08:00"))  # trop vieux
        ws.append(row("SELL", "MOM", "x", None, None, "Ouvert"))
        wb.save(os.path.join(self.dir, "trades.xlsx"))
        out = []
        with mock.patch.object(S, "ROOT", self.dir), \
             mock.patch("builtins.print", lambda *a, **k: out.append(" ".join(map(str, a)))):
            S.ghost_report("2026-10-27 16:09")
        txt = "\n".join(out)
        self.assertIn("clos depuis 2026-10-27 16:09 UTC : 2 · encore ouverts : 1", txt)
        self.assertIn("aurait été tradé (stratégie actuelle) :  n=   1  WR=100%  moy=+0.500%", txt)
        self.assertIn("filtrés / bloqués :  n=   1  WR=  0%  moy=-0.200%", txt)

    def test_bilan_tendance(self):
        with open(os.path.join(self.dir, "trend_state.json"), "w") as f:
            json.dump({"positions": {"ETH": {"side": "BUY", "entry": 2000, "sl": 1800, "paper": True,
                                             "opened_at": 0}}}, f)
        wb = Workbook(); wb.active.title = "Trades"
        ws = wb.create_sheet("Tendance")
        ws.append(["ID", "Ouvert (UTC)", "Symbole", "Côté", "Mode", "Qté", "Entrée", "SL",
                   "Notionnel $", "Statut", "Sortie", "P&L $", "P&L %", "Fermé (UTC)"])
        row = lambda mode, pnl, pct: ["P", "", "SOL/USDT", "SELL", mode, 1, 1, 1, 11, "x", 1, pnl, pct, ""]
        ws.append(row("Fantôme", 1.5, 13.0))
        ws.append(row("Fantôme", -0.5, -5.0))
        ws.append(row("Réel", None, None))        # encore ouverte
        wb.save(os.path.join(self.dir, "trades.xlsx"))
        out = []
        with mock.patch.object(S, "ROOT", self.dir), \
             mock.patch("builtins.print", lambda *a, **k: out.append(" ".join(map(str, a)))):
            S.trend_report()
        txt = "\n".join(out)
        self.assertIn("Positions en cours : 1", txt)
        self.assertIn("ETH   LONG  entrée 2000", txt)
        self.assertIn("fantôme", txt)
        self.assertIn("Fermées (fantôme) : 2 · gagnantes 1 · P&L $+1.00 · moy +4.00 % du notionnel", txt)
        self.assertNotIn("Fermées (réel)", txt)


if __name__ == "__main__":
    unittest.main()
