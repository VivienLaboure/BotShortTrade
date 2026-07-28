"""Test rapide : connexion XTB → achat BITCOIN minimum → revente immédiate."""
import os, json, time
import websocket as _ws
from dotenv import load_dotenv

load_dotenv()
USER_ID  = os.getenv("XTB_USER_ID")
PASSWORD = os.getenv("XTB_PASSWORD")
URL      = "wss://ws.xtb.com/demo"

def send(ws, cmd):
    ws.send(json.dumps(cmd))
    return json.loads(ws.recv())

print("1. Connexion à XTB demo...")
ws = _ws.create_connection(URL, timeout=15)

print("2. Login...")
r = send(ws, {"command": "login", "arguments": {"userId": USER_ID, "password": PASSWORD}})
if not r.get("status"):
    print(f"   ❌ Login échoué: {r}")
    ws.close()
    exit(1)
print(f"   ✅ Connecté (streamSessionId: {r.get('streamSessionId', 'N/A')[:20]}...)")

print("3. Info symbole BITCOIN...")
r = send(ws, {"command": "getSymbol", "arguments": {"symbol": "BITCOIN"}})
if not r.get("status"):
    print(f"   ❌ Symbole BITCOIN non trouvé: {r}")
    print("   → Essai avec BTCUSD...")
    r = send(ws, {"command": "getSymbol", "arguments": {"symbol": "BTCUSD"}})
    if not r.get("status"):
        print(f"   ❌ BTCUSD non trouvé non plus: {r}")
        ws.close()
        exit(1)
    symbol = "BTCUSD"
else:
    symbol = "BITCOIN"

info = r["returnData"]
digits   = info.get("digits", 2)
lot_min  = info.get("lotMin", 0.01)
lot_step = info.get("lotStep", 0.01)
cs       = info.get("contractSize", 1)
print(f"   ✅ {symbol} | digits={digits} | lotMin={lot_min} | contractSize={cs}")

print("4. Prix actuel...")
r = send(ws, {"command": "getTickPrices",
              "arguments": {"symbols": [symbol], "timestamp": 0, "level": 0}})
if not r.get("status") or not r["returnData"].get("quotations"):
    print(f"   ❌ Prix indisponible: {r}")
    ws.close()
    exit(1)
q = r["returnData"]["quotations"][0]
ask, bid = float(q["ask"]), float(q["bid"])
spread = round(ask - bid, digits)
print(f"   ✅ Ask=${ask:.2f} | Bid=${bid:.2f} | Spread=${spread}")

print(f"5. Ouverture BUY {lot_min} lot {symbol} @ ${ask:.2f}...")
r = send(ws, {
    "command": "tradeTransaction",
    "arguments": {"tradeTransInfo": {
        "cmd": 0,            # BUY
        "symbol": symbol,
        "volume": lot_min,
        "price": ask,
        "sl": round(ask * 0.995, digits),   # SL -0.5%
        "tp": round(ask * 1.005, digits),   # TP +0.5%
        "type": 0,           # OPEN
        "comment": "test-scalper",
        "expiration": 0, "offset": 0, "order": 0,
    }}
})
if not r.get("status"):
    print(f"   ❌ Ordre refusé: {r}")
    ws.close()
    exit(1)
order_ref = r["returnData"]["order"]
print(f"   ✅ Ordre envoyé (ref={order_ref})")

print("6. Vérification statut...")
time.sleep(1)
r = send(ws, {"command": "tradeTransactionStatus", "arguments": {"order": order_ref}})
rd = r.get("returnData", {})
req_status = rd.get("requestStatus")
order_id   = rd.get("order2", order_ref)
status_map = {0: "erreur", 1: "pending", 3: "accepté", 4: "rejeté"}
print(f"   Status: {status_map.get(req_status, req_status)} | order2={order_id}")
if req_status not in (1, 3):
    print(f"   ❌ Ordre rejeté: {rd.get('message', rd)}")
    ws.close()
    exit(1)
print(f"   ✅ Position ouverte (ID: {order_id})")

print("7. Attente 3 secondes...")
time.sleep(3)

print("8. Vérification position ouverte...")
r = send(ws, {"command": "getTrades", "arguments": {"openedOnly": True}})
trades = r.get("returnData", [])
our = [t for t in trades if t.get("symbol") == symbol and "test-scalper" in str(t.get("comment",""))]
if not our:
    print(f"   ⚠️  Position introuvable dans getTrades (peut être déjà fermée ?)")
else:
    t = our[0]
    print(f"   ✅ Position active | profit flottant: ${t.get('profit', 0):.2f}")

print(f"9. Fermeture de la position (ID={order_id})...")
r = send(ws, {"command": "getTickPrices",
              "arguments": {"symbols": [symbol], "timestamp": 0, "level": 0}})
bid_close = float(r["returnData"]["quotations"][0]["bid"])

r = send(ws, {
    "command": "tradeTransaction",
    "arguments": {"tradeTransInfo": {
        "cmd": 1,            # SELL (ferme un BUY)
        "symbol": symbol,
        "volume": lot_min,
        "price": bid_close,
        "sl": 0, "tp": 0,
        "type": 2,           # CLOSE
        "order": int(order_id),
        "comment": "test-close",
        "expiration": 0, "offset": 0,
    }}
})
if not r.get("status"):
    print(f"   ❌ Fermeture échouée: {r}")
else:
    print(f"   ✅ Position fermée @ ${bid_close:.2f}")

print("\n✅ TEST COMPLET — XTB fonctionne correctement !")
ws.close()
