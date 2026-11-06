"""Réglages du serveur, lus dans l'environnement (fichier .env en production).

La clé Anthropic n'apparaît nulle part dans le code : le SDK la lit dans
ANTHROPIC_API_KEY. Elle ne quitte jamais le serveur.
"""

import os
from pathlib import Path


def _env(nom, defaut):
    v = os.environ.get(nom)
    return v if v not in (None, "") else defaut


MODELE = _env("SMARTMAIL_MODELE", "claude-opus-5-5")
# Le tri est une tâche de classement : un effort bas suffit (moins de jetons de réflexion).
EFFORT = _env("SMARTMAIL_EFFORT", "low")
TAILLE_LOT = int(_env("SMARTMAIL_TAILLE_LOT", "150"))
PARALLELE = int(_env("SMARTMAIL_PARALLELE", "4"))  # lots en parallèle pour UNE demande
APPELS_MAX = int(_env("SMARTMAIL_APPELS_MAX", "8"))  # appels simultanés à Claude pour TOUT le serveur
# Plafond de dépense mensuel, tous utilisateurs confondus (en dollars). Fixez aussi
# une limite dans la console Anthropic : c'est la vraie ceinture de sécurité.
BUDGET_MENSUEL_USD = float(_env("SMARTMAIL_BUDGET_MENSUEL_USD", "50"))
MAX_CONTACTS_PAR_REQUETE = int(_env("SMARTMAIL_MAX_CONTACTS_PAR_REQUETE", "2000"))
CACHE_JOURS = int(_env("SMARTMAIL_CACHE_JOURS", "180"))
BASE = Path(_env("SMARTMAIL_BASE", str(Path(__file__).resolve().parent.parent / "data" / "smartmail.db")))
# Mode démonstration / tests : un faux classeur déterministe remplace Claude.
FAUX_CLAUDE = _env("SMARTMAIL_FAUX_CLAUDE", "0") == "1"
DEMO = _env("SMARTMAIL_DEMO", "0") == "1"  # faux annuaire + accès de démonstration (/demo/...)
FAUX_LATENCE = float(_env("SMARTMAIL_FAUX_LATENCE", "0"))  # secondes par lot, en simulation
ORIGINES = [o.strip() for o in _env(
    "SMARTMAIL_ORIGINES",
    "https://outlook.office.com,https://outlook.office365.com,https://outlook.live.com,https://outlook.cloud.microsoft",
).split(",") if o.strip()]

# Prix par million de jetons (entrée, sortie), en dollars : sert au suivi des
# dépenses et au plafond. À tenir à jour si vous changez de modèle.
PRIX = {
    "claude-opus-5-5": (4.00, 20.00),
    "claude-sonnet-5-5": (2.00, 10.00),
    "claude-haiku-5-5": (0.10, 0.50),
    "claude-opus-4-8": (5.00, 25.00),
}


def cout(modele, jetons_entree, jetons_sortie):
    pe, ps = PRIX.get(modele, PRIX["claude-opus-5-5"])
    return (jetons_entree * pe + jetons_sortie * ps) / 1_000_000
