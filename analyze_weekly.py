"""
analyze_weekly.py — Analyse hebdomadaire des trades (statistiques locales)

Usage :
    python analyze_weekly.py

Lit trades.xlsx + lessons.json et affiche un rapport de performance complet
sans appel à une API externe.

Pré-requis :
    pip install openpyxl python-dotenv
"""
import json
import os
from collections import defaultdict
from datetime import datetime, timezone, timedelta

import openpyxl
from dotenv import load_dotenv

load_dotenv()

EXCEL_FILE   = "trades.xlsx"
LESSONS_FILE = "lessons.json"


# ── Lecture des trades ─────────────────────────────────────────────────────────
def load_trades() -> list[dict]:
    if not os.path.exists(EXCEL_FILE):
        print(f"[ERREUR] {EXCEL_FILE} introuvable — lance d'abord le bot.")
        return []
    wb   = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
    ws   = wb["Trades"]
    hdrs = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[14] is None:       # "P&L final $" vide = trade encore ouvert
            continue
        rows.append(dict(zip(hdrs, row)))
    wb.close()
    return rows


# ── Calcul des stats ───────────────────────────────────────────────────────────
def compute_stats(trades: list[dict]) -> dict:
    cutoff_7d  = datetime.now(timezone.utc) - timedelta(days=7)
    cutoff_30d = datetime.now(timezone.utc) - timedelta(days=30)

    stats: dict = {
        "total_trades": len(trades),
        "by_symbol":   defaultdict(lambda: {"w": 0, "l": 0, "pnl": 0.0}),
        "by_hour":     defaultdict(lambda: {"w": 0, "l": 0, "pnl": 0.0}),
        "by_outcome":  {"TP": 0, "SL": 0, "other": 0},
        "week_trades": 0,
        "week_pnl":    0.0,
        "month_trades": 0,
        "month_pnl":    0.0,
        "gross_profit": 0.0,
        "gross_loss":   0.0,
    }

    wins, losses = [], []

    for t in trades:
        pnl    = float(t.get("P&L final $", 0) or 0)
        coin   = str(t.get("Symbole", "?")).replace("/USDT", "")
        status = str(t.get("Statut", ""))
        date   = t.get("Date/Heure")

        # Heure d'ouverture
        hour = "?"
        if date:
            try:
                if isinstance(date, str):
                    dt = datetime.fromisoformat(date.replace(" ", "T"))
                else:
                    dt = date  # déjà un datetime (openpyxl)
                hour = str(dt.hour)
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                if dt >= cutoff_7d:
                    stats["week_trades"] += 1
                    stats["week_pnl"]    += pnl
                if dt >= cutoff_30d:
                    stats["month_trades"] += 1
                    stats["month_pnl"]    += pnl
            except Exception:
                pass

        # Par symbole
        s = stats["by_symbol"][coin]
        if pnl > 0:
            s["w"]   += 1
            s["pnl"] += pnl
            stats["gross_profit"] += pnl
            wins.append(pnl)
        else:
            s["l"]   += 1
            s["pnl"] += pnl
            stats["gross_loss"] += abs(pnl)
            losses.append(pnl)

        # Par heure
        h = stats["by_hour"][hour]
        h["w" if pnl > 0 else "l"] += 1
        h["pnl"] += pnl

        # Par outcome (TP / SL)
        su = status.upper()
        if "TP" in su:
            stats["by_outcome"]["TP"] += 1
        elif "SL" in su:
            stats["by_outcome"]["SL"] += 1
        else:
            stats["by_outcome"]["other"] += 1

    stats["avg_win"]  = sum(wins)   / len(wins)   if wins   else 0.0
    stats["avg_loss"] = sum(losses) / len(losses) if losses else 0.0

    # Convertir defaultdict → dict normal
    stats["by_symbol"] = dict(stats["by_symbol"])
    stats["by_hour"]   = dict(stats["by_hour"])

    return stats


# ── Affichage formaté ──────────────────────────────────────────────────────────
def print_report(stats: dict, lessons: dict):
    tp  = stats["by_outcome"]["TP"]
    sl  = stats["by_outcome"]["SL"]
    n   = stats["total_trades"]
    wr  = round(tp / n * 100, 1) if n > 0 else 0
    pf  = round(stats["gross_profit"] / stats["gross_loss"], 2) if stats["gross_loss"] > 0 else 0
    net = stats["gross_profit"] - stats["gross_loss"]

    w = 60
    sep  = "-" * w
    sep2 = "=" * w

    print(f"\n{sep2}")
    print(f"  RAPPORT DE PERFORMANCE - BotShortTrade")
    print(f"  Genere le {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(sep2)

    print("\nGLOBAL")
    print(f"  Total trades  : {n}")
    print(f"  TP / SL       : {tp} / {sl}  ({stats['by_outcome']['other']} autres)")
    print(f"  Win Rate      : {wr}%")
    print(f"  Profit Factor : {pf}")
    print(f"  P&L net       : ${net:+.2f}")
    print(f"  Gain moyen    : +${stats['avg_win']:.2f}")
    print(f"  Perte moyenne : -${abs(stats['avg_loss']):.2f}")

    print(f"\n{sep}")
    print(f"  7 DERNIERS JOURS")
    print(sep)
    w7 = stats["week_trades"]
    pnl7 = stats["week_pnl"]
    print(f"  Trades  : {w7}")
    print(f"  P&L     : ${pnl7:+.2f}")

    print(f"\n{sep}")
    print(f"  30 DERNIERS JOURS")
    print(sep)
    w30 = stats["month_trades"]
    pnl30 = stats["month_pnl"]
    print(f"  Trades  : {w30}")
    print(f"  P&L     : ${pnl30:+.2f}")

    # ── Par symbole ──
    print(f"\n{sep}")
    print(f"  PAR SYMBOLE")
    print(sep)
    print(f"  {'Crypto':<8}  {'W':>4}  {'L':>4}  {'WR':>6}  {'P&L':>8}")
    for sym, v in sorted(stats["by_symbol"].items(), key=lambda x: x[1]["pnl"]):
        tot = v["w"] + v["l"]
        wr_s = round(v["w"] / tot * 100) if tot else 0
        flag = "!" if (wr_s < 35 and tot >= 5) else ("+" if wr_s >= 55 else " ")
        print(f"  [{flag}] {sym:<7}  {v['w']:>4}  {v['l']:>4}  {wr_s:>5}%  {v['pnl']:>+7.2f}$")

    # ── Par heure ──
    print(f"\n{sep}")
    print(f"  PAR HEURE UTC")
    print(sep)
    print(f"  {'H':>3}  {'W':>4}  {'L':>4}  {'WR':>6}  {'P&L':>8}")
    for h, v in sorted(stats["by_hour"].items(), key=lambda x: int(x[0]) if x[0].isdigit() else 99):
        tot = v["w"] + v["l"]
        wr_h = round(v["w"] / tot * 100) if tot else 0
        flag = " OK" if wr_h >= 60 else ("BAD" if wr_h < 30 and tot >= 3 else "   ")
        print(f"  {flag} {h:>2}h  {v['w']:>4}  {v['l']:>4}  {wr_h:>5}%  {v['pnl']:>+7.2f}$")

    # ── Leçons ──
    buckets = lessons.get("buckets", {})
    if buckets:
        print(f"\n{sep}")
        print(f"  PATTERNS LES PLUS PERDANTS (top 10)")
        print(sep)
        worst = sorted(buckets.items(), key=lambda x: x[1].get("pnl", 0))[:10]
        for k, v in worst:
            tot  = v.get("w", 0) + v.get("l", 0)
            wr_l = round(v.get("w", 0) / tot * 100) if tot else 0
            print(f"  {k:<35}  WR={wr_l:>3}%  {v.get('w',0)}W/{v.get('l',0)}L  {v.get('pnl',0):>+7.2f}$")

    print(f"\n{sep2}\n")


# ── Sauvegarde du rapport ──────────────────────────────────────────────────────
def save_report(stats: dict, lessons: dict) -> str:
    tp  = stats["by_outcome"]["TP"]
    sl  = stats["by_outcome"]["SL"]
    n   = stats["total_trades"]
    wr  = round(tp / n * 100, 1) if n > 0 else 0
    pf  = round(stats["gross_profit"] / stats["gross_loss"], 2) if stats["gross_loss"] > 0 else 0
    net = stats["gross_profit"] - stats["gross_loss"]
    ts  = datetime.now().strftime("%Y-%m-%d_%H%M")
    path = f"report_weekly_{ts}.md"

    lines = [
        f"# Rapport de performance — {datetime.now().strftime('%Y-%m-%d %H:%M')}\n",
        f"## Résumé global\n",
        f"| Indicateur | Valeur |",
        f"|---|---|",
        f"| Total trades | {n} |",
        f"| TP / SL | {tp} / {sl} |",
        f"| Win Rate | {wr}% |",
        f"| Profit Factor | {pf} |",
        f"| P&L net | ${net:+.2f} |",
        f"| Gain moyen | +${stats['avg_win']:.2f} |",
        f"| Perte moyenne | -${abs(stats['avg_loss']):.2f} |",
        f"\n## 7 derniers jours\n",
        f"- Trades : {stats['week_trades']}",
        f"- P&L : ${stats['week_pnl']:+.2f}",
        f"\n## Performance par symbole\n",
        f"| Crypto | W | L | WR | P&L |",
        f"|---|---|---|---|---|",
    ]
    for sym, v in sorted(stats["by_symbol"].items(), key=lambda x: x[1]["pnl"]):
        tot = v["w"] + v["l"]
        wr_s = round(v["w"] / tot * 100) if tot else 0
        lines.append(f"| {sym} | {v['w']} | {v['l']} | {wr_s}% | ${v['pnl']:+.2f} |")

    lines += [f"\n## Performance par heure UTC\n", f"| H | W | L | WR | P&L |", f"|---|---|---|---|---|"]
    for h, v in sorted(stats["by_hour"].items(), key=lambda x: int(x[0]) if x[0].isdigit() else 99):
        tot = v["w"] + v["l"]
        wr_h = round(v["w"] / tot * 100) if tot else 0
        lines.append(f"| {h}h | {v['w']} | {v['l']} | {wr_h}% | ${v['pnl']:+.2f} |")

    buckets = lessons.get("buckets", {})
    if buckets:
        lines += [f"\n## Patterns les plus perdants\n", f"| Pattern | WR | W/L | P&L |", f"|---|---|---|---|"]
        for k, v in sorted(buckets.items(), key=lambda x: x[1].get("pnl", 0))[:15]:
            tot  = v.get("w", 0) + v.get("l", 0)
            wr_l = round(v.get("w", 0) / tot * 100) if tot else 0
            lines.append(f"| {k} | {wr_l}% | {v.get('w',0)}W/{v.get('l',0)}L | ${v.get('pnl',0):+.2f} |")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  BotShortTrade — Analyse Hebdomadaire")
    print("=" * 60)

    print("\n  Chargement des trades…")
    trades = load_trades()
    if not trades:
        return
    print(f"  {len(trades)} trades chargés depuis {EXCEL_FILE}")

    print("  Chargement des leçons…")
    lessons = {}
    if os.path.exists(LESSONS_FILE):
        with open(LESSONS_FILE, encoding="utf-8") as f:
            lessons = json.load(f)
        print(f"  {len(lessons.get('buckets', {}))} patterns chargés depuis {LESSONS_FILE}")

    print("  Calcul des statistiques…")
    stats = compute_stats(trades)

    print_report(stats, lessons)

    path = save_report(stats, lessons)
    print(f"  [OK] Rapport sauvegarde : {path}\n")


if __name__ == "__main__":
    main()
