"""Évaluation avec le VRAI Claude sur les 8 écoles fictives (app/ecoles.py).

Pour chaque école : échantillon de 30 (comme l'app) → Claude apprend la logique →
tri complet (règles → logique → Claude) de N contacts → intrus, rappel, coût.

    ANTHROPIC_API_KEY=... uv run python -m outils.eval_ecoles --contacts 600 --oui

Sans --oui, affiche seulement l'estimation du coût. Base SQLite jetable (aucune
trace dans la base de production).
"""

import argparse
import sys
import tempfile
from pathlib import Path

from app import base, config, ecoles, service
from app.ia import ClasseurClaude
from app.logique import ApprentiClaude

PIEGES = {"service", "association", "perso", "externe", "invite", "salle", "liste"}


def echantillon(page):
    avec = [c for c in page if c["intitule"]]
    sans = [c for c in page if not c["intitule"]]
    return (avec[:22] + sans[:30 - min(22, len(avec))])[:30]


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("--contacts", type=int, default=600, help="contacts triés par école")
    p.add_argument("--ecoles", default=",".join(ecoles.PROFILS))
    p.add_argument("--oui", action="store_true", help="lancer pour de vrai (consomme des crédits Anthropic)")
    a = p.parse_args(argv)
    liste = [e for e in a.ecoles.split(",") if e in ecoles.PROFILS]

    # Estimation prudente : ~45 jetons d'entrée et ~35 de sortie (réflexion comprise) par contact envoyé à Claude.
    pe, ps = config.PRIX.get(config.MODELE, (4, 20))
    estimation = len(liste) * (a.contacts * (45 * pe + 35 * ps) / 1e6 + (3000 * pe + 2000 * ps) / 1e6)
    print(f"Modèle {config.MODELE} · {len(liste)} écoles × {a.contacts} contacts · coût maximal estimé ≈ {estimation:.2f} $")
    if not a.oui:
        print("Ajoutez --oui pour lancer.")
        return 0

    config.BASE = Path(tempfile.mkdtemp()) / "eval.db"
    config.BUDGET_MENSUEL_USD = max(5.0, estimation * 2)
    base.fermer()
    service.definir_classeur(ClasseurClaude())
    apprenti = ApprentiClaude()
    jeton = base.creer_utilisateur("evaluation", 10**7)
    u = base.utilisateur_par_jeton(jeton)

    total_cout, echecs = 0.0, 0
    print(f"\n{'École':16} {'Intrus':>6} {'Rappel':>7} {'Incert.':>7} {'Coût $':>7}  Logique apprise")
    for nom in liste:
        annuaire = ecoles.PROFILS[nom]()
        lg, cout_lg, err = apprenti.apprendre(echantillon(annuaire[:200]))
        if err:
            print(f"{nom:16} logique : {err}")
        lot = annuaire[:a.contacts]
        r = service.trier(u, lot, lg)
        verite = {c["adresse"].lower(): c["verite"] for c in lot}
        retenus = [x for x in r["resultats"] if x["groupe"] == "etudiant"]
        intrus = [x for x in retenus if verite[x["adresse"]] != "etudiant"]
        vrais = sum(1 for v in verite.values() if v == "etudiant")
        incertains = sum(1 for x in r["resultats"] if x["groupe"] == "incertain")
        cout = cout_lg + r["cout_usd"]
        total_cout += cout
        echecs += bool(intrus)
        print(f"{nom:16} {len(intrus):>6} {len(retenus) / max(1, vrais):>7.0%} {incertains:>7} {cout:>7.3f}  {(lg.resume if lg else '—')[:70]}")
        for x in intrus[:5]:
            print(f"   ⚠ intrus : {x['adresse']} ({verite[x['adresse']]}) — {x['source']} : {x['raison']}")
        if r["erreurs"]:
            print(f"   erreurs : {r['erreurs']}")
    print(f"\nCoût total : {total_cout:.2f} $ · écoles avec intrus : {echecs}/{len(liste)}")
    return 1 if echecs else 0


if __name__ == "__main__":
    sys.exit(main())
