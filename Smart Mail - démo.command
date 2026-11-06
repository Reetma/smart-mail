#!/bin/zsh
# Double-clic : lance Smart Mail en démonstration et l'ouvre dans le navigateur.
# (Faux annuaire d'école, faux Claude, aucun mail envoyé.) Fermez cette fenêtre pour arrêter.
cd "$(dirname "$0")/serveur" || exit 1
export SMARTMAIL_DEMO=1 SMARTMAIL_FAUX_CLAUDE=1 SMARTMAIL_FAUX_LATENCE=1 SMARTMAIL_BASE="$HOME/Library/Application Support/SmartMail-demo/demo.db"
mkdir -p "$HOME/Library/Application Support/SmartMail-demo"
PORT=8766
if curl -sf "http://127.0.0.1:$PORT/sante" >/dev/null; then open "http://127.0.0.1:$PORT"; exit 0; fi
( for i in {1..40}; do curl -sf "http://127.0.0.1:$PORT/sante" >/dev/null && open "http://127.0.0.1:$PORT" && break; sleep 0.5; done ) &
echo "Smart Mail (démo) sur http://127.0.0.1:$PORT — fermez cette fenêtre pour arrêter."
exec "$HOME/.local/bin/uv" run uvicorn app.main:app --host 127.0.0.1 --port $PORT --log-level warning
