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


if __name__ == "__main__":
    unittest.main()
