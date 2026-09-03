#!/usr/bin/env python3
"""
reset_history.py — Efface l'historique de trading sans perturber les positions ouvertes.

Ce script :
  1. Archive puis supprime trades.xlsx  (historique Excel)
  2. Archive puis supprime lessons.json (patterns appris)
  3. Réécrit bot_state.json en conservant open_positions
     mais en supprimant initial_balance → le bot recalibrera sa
     référence P&L sur le solde réel au prochain démarrage.

Utilisation :
  python reset_history.py
"""
import json, os, shutil
from datetime import datetime, timezone

STATE_FILE   = "bot_state.json"
EXCEL_FILE   = "trades.xlsx"
LESSONS_FILE = "lessons.json"

now_str = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
print("=" * 55)
print("  Reset historique BotShortTrade")
print("=" * 55)

# ── 1. Archiver et supprimer trades.xlsx ──────────────────────
if os.path.exists(EXCEL_FILE):
    archive = f"trades_archive_{now_str}.xlsx"
    shutil.copy(EXCEL_FILE, archive)
    os.remove(EXCEL_FILE)
    print(f"✓ trades.xlsx    → archivé : {archive}")
else:
    print("  trades.xlsx      introuvable — ignoré")

# Nettoyer les .tmp résiduels
for tmp in [EXCEL_FILE + ".tmp", STATE_FILE + ".tmp"]:
    if os.path.exists(tmp):
        os.remove(tmp)
        print(f"✓ {tmp} supprimé")

# ── 2. Archiver et supprimer lessons.json ────────────────────
if os.path.exists(LESSONS_FILE):
    archive_l = f"lessons_archive_{now_str}.json"
    shutil.copy(LESSONS_FILE, archive_l)
    os.remove(LESSONS_FILE)
    print(f"✓ lessons.json   → archivé : {archive_l}")
else:
    print("  lessons.json     introuvable — ignoré")

# ── 3. Préserver open_positions, effacer initial_balance ──────
open_positions = []
if os.path.exists(STATE_FILE):
    try:
        with open(STATE_FILE) as f:
            state = json.load(f)
        open_positions = state.get("open_positions", [])
    except Exception as e:
        print(f"  [WARN] lecture {STATE_FILE}: {e}")

# Réécrire avec seulement open_positions
new_state = {"open_positions": open_positions}
tmp = STATE_FILE + ".tmp"
with open(tmp, "w") as f:
    json.dump(new_state, f, indent=2)
os.replace(tmp, STATE_FILE)
print(f"✓ bot_state.json → open_positions conservées, initial_balance réinitialisé")

# ── Résumé ────────────────────────────────────────────────────
print()
print("Au prochain démarrage du bot :")
print("  • Nouveau trades.xlsx créé automatiquement")
print("  • initial_balance = solde réel actuel (P&L remis à 0)")
print("  • lessons.json recréé vide (apprentissage repart de zéro)")
if open_positions:
    coins = [p.get("symbol", "?") + " " + p.get("side", "?")
             for p in open_positions]
    print(f"  • {len(open_positions)} position(s) reprise(s) : {', '.join(coins)}")
else:
    print("  • Aucune position ouverte à reprendre")
print("=" * 55)
