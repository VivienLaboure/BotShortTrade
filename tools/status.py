"""
État de BotShortTrade en une commande, sans copier-coller Discord :
  python3 tools/status.py [--hours 24]

Lit uniquement des fichiers locaux et le journal systemd (aucun secret, aucun appel réseau) :
  service systemd, erreurs du journal, messages Discord (discord_events.jsonl),
  stratégie tendance (trend_state.json + onglet Tendance de trades.xlsx)
"""
from __future__ import annotations
import argparse, json, os, re, subprocess
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ANSI = re.compile(r"\x1b\[[0-9;]*m")


def sh(cmd: list[str]) -> str:
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=30).stdout
    except Exception as e:
        return f"({e})"


def read_events(path: str, since: datetime) -> list[dict]:
    """Messages Discord journalisés depuis `since` (lignes JSON illisibles ignorées)."""
    out = []
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8") as f:
        for line in f:
            try:
                ev = json.loads(line)
                if datetime.fromisoformat(ev["t"]) >= since:
                    out.append(ev)
            except Exception:
                continue
    return out


def print_events(events: list[dict], full_prefix: str):
    """Un message par ligne ; le dernier dont le titre commence par `full_prefix` en entier."""
    if not events:
        print("  (aucun message — le journal local n'existe que depuis le redémarrage qui l'a ajouté)")
        return
    for ev in events:
        first = (ev.get("description") or "").split("\n")[0]
        print(f"  {ev['t'][5:16]}  {ev['title']} — {first[:110]}")
    last = [ev for ev in events if ev["title"].startswith(full_prefix)]
    if last:
        ev = last[-1]
        print(f"\n  Dernier « {ev['title']} » ({ev['t'][5:16]}) :")
        for line in (ev.get("description") or "").split("\n"):
            print(f"    {line}")
        for k, v in (ev.get("fields") or {}).items():
            print(f"    {k} : {v}")


def trend_report():
    """Positions tendance en cours (trend_state.json) et bilan des positions fermées (onglet Tendance)."""
    try:
        with open(os.path.join(ROOT, "trend_state.json"), encoding="utf-8") as f:
            positions = json.load(f).get("positions", {})
    except Exception:
        positions = {}
    now = datetime.now(timezone.utc).timestamp()
    print(f"  Positions en cours : {len(positions)}")
    for coin, p in sorted(positions.items()):
        print(f"    {coin:5s} {'LONG ' if p['side'] == 'BUY' else 'SHORT'} entrée {p['entry']:<10.6g} "
              f"SL {p['sl']:<10.6g} {'fantôme' if p.get('paper') else 'réel'} · "
              f"depuis {(now - p.get('opened_at', now)) / 86400:.1f} j")
    path = os.path.join(ROOT, "trades.xlsx")
    if not os.path.exists(path):
        return
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    if "Tendance" not in wb.sheetnames:
        return
    rows = list(wb["Tendance"].iter_rows(values_only=True))
    h = {name: i for i, name in enumerate(rows[0])}
    closed = [r for r in rows[1:] if r[h["P&L $"]] is not None]
    for mode in ("Réel", "Fantôme"):
        rs = [r for r in closed if r[h["Mode"]] == mode]
        if rs:
            pnl = [r[h["P&L $"]] for r in rs]
            print(f"  Fermées ({mode.lower()}) : {len(rs)} · gagnantes {sum(x > 0 for x in pnl)} · "
                  f"P&L ${sum(pnl):+.2f} · moy {sum(r[h['P&L %']] for r in rs) / len(rs):+.2f} % du notionnel")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--hours", type=float, default=24, help="fenêtre du journal et des messages (h)")
    a = ap.parse_args()
    since = datetime.now(timezone.utc) - timedelta(hours=a.hours)

    print(f"═══ BotShortTrade ═══ (dernières {a.hours:g} h)")
    state = sh(["systemctl", "--user", "show", "botshorttrade", "-p", "ActiveState,ActiveEnterTimestamp"])
    print("  Service : " + " · ".join(state.strip().splitlines()))
    journal = ANSI.sub("", sh(["journalctl", "--user", "-u", "botshorttrade", "--since", f"-{int(a.hours * 60)}min",
                               "--no-pager", "-o", "cat"]))
    errs: dict[str, int] = {}
    for line in journal.splitlines():
        if re.search(r"ERR|WARN|Traceback|Exception", line):
            key = re.sub(r"\d\d:\d\d:\d\d", "", line.strip())
            errs[key] = errs.get(key, 0) + 1
    print(f"  Journal : {len(journal.splitlines())} lignes · {sum(errs.values())} erreurs/alertes")
    for k, n in sorted(errs.items(), key=lambda kv: -kv[1])[:10]:
        print(f"    ×{n:<3d} {k[:140]}")
    print("\n  Messages Discord :")
    print_events(read_events(os.path.join(ROOT, "discord_events.jsonl"), since), "💓")
    print("\n  Stratégie tendance :")
    trend_report()


if __name__ == "__main__":
    main()
