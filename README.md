# BotShortTrade

Bot de trading sur Hyperliquid (perpétuels), stratégie de **tendance journalière** :

- chaque jour à 00:00 UTC, pour BTC, ETH, SOL, XRP et DOGE : LONG si la clôture du jour dépasse celle
  d'il y a 14, 28 et 56 jours pour au moins 2 de ces horizons, SHORT sinon ; position tenue jusqu'à
  l'inversion du signal ;
- stop-loss posé sur Hyperliquid à 3 × ATR14 journalier ; après un SL, à plat jusqu'au lendemain ;
- taille : `TREND_EXPOSURE` (50 %) de l'équité répartie sur les 5 coins, aucun ordre sous 10,5 $.

Résultats des tests sur le passé et variantes rejetées : docstring de `tools/trend_backtest.py`.
L'ancien scalping M1 (aucun avantage après frais) a été retiré le 2026-11-11 ; il reste dans
l'historique git (étiquette `scalping-final`).

## Fichiers

| Fichier | Rôle |
|---|---|
| `main.py` | le bot : stratégie tendance, ordres, alertes Discord, boucle d'une minute |
| `tools/status.py` | état en une commande : service, erreurs du journal, messages Discord, positions |
| `tools/trend_backtest.py` | test de la règle sur l'historique (données publiques Binance et Bybit) |
| `tests/` | tests de non-régression, dont un Hyperliquid simulé sur des mois de prix aléatoires |
| `.env` | clé privée et réglages (jamais commité ; modèle : `.env.example`) |
| `trend_state.json` | positions suivies, SL, ordres restés sans réponse |
| `bot_state.json` | référence du « % depuis le départ » et dernier relevé des dépôts |
| `trades.xlsx` | une ligne par position (onglet Tendance) |
| `discord_events.jsonl` | copie locale de chaque message Discord |
| `bot_errors.log` | copie datée de chaque erreur (le journal systemd ne garde qu'environ 1,5 jour) |
| `dashboard.html` | tableau de bord local, régénéré à chaque cycle |

Les fichiers de données et `.env` sont exclus de git (`.gitignore`).

## Exploitation

Le bot tourne en service systemd utilisateur, relancé automatiquement en cas d'erreur :

```
systemctl --user status botshorttrade
systemctl --user restart botshorttrade     # après toute modification de main.py
journalctl --user -u botshorttrade -f      # journal en direct
python3 tools/status.py                    # résumé des dernières 24 h
```

Ne jamais lancer `python3 main.py` à la main pendant que le service tourne : deux bots sur le même
compte. Le compte est partagé avec un autre bot (BotCopyTrade) : BotShortTrade n'ouvre jamais un coin
déjà tenu sur le compte par une position qui n'est pas la sienne.

## Modifier le code

```
pip install -r requirements.txt            # versions figées et vérifiées
python3 -m unittest discover -s tests      # avant chaque commit ; aucun réseau, aucun ordre réel
```

Chaque correctif ajoute un test qui échoue sur l'ancien code.
