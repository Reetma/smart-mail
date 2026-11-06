#!/bin/sh
# Envoie Smart Mail sur le VPS et (re)démarre tout.  Usage : ./deploy.sh ubuntu@51.83.12.34
set -e
CIBLE="$1"; [ -n "$CIBLE" ] || { echo "Usage : $0 utilisateur@ip-du-vps"; exit 1; }
ICI="$(cd "$(dirname "$0")/.." && pwd)"
rsync -az --delete --exclude .venv --exclude data --exclude __pycache__ --exclude .pytest_cache --exclude 'deploiement/.env' "$ICI/" "$CIBLE:smartmail/"
ssh "$CIBLE" 'cd smartmail/deploiement && test -f .env || { echo "Créez smartmail/deploiement/.env (voir .env.exemple)"; exit 1; } && docker compose up -d --build && docker compose ps'
