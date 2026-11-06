"""Banc d'essai sur des écoles aux conventions différentes (app/ecoles.py).

Règle d'or : AUCUN non-étudiant ne doit être retenu (professeur, service, salle,
liste, invité, contact perso, partenaire…), quelle que soit l'école. On mesure
ensuite la part des vrais étudiants retrouvés.
"""

import pytest

from app import ecoles, logique, regles, service
from app.ia import Lot
from app.logique import FauxApprenti, Logique, RegleDomaine, format_adresse

PIEGES = {"service", "association", "perso", "externe", "invite", "salle", "liste"}
MOTS_E = ["etudiant", "etudiante", "student", "undergraduate", "msc", "pge", "bachelor", "programme"]
MOTS_P = ["professeur", "professeure", "maitre", "assistante", "responsable", "chargee", "directeur", "enseignant",
          "gestionnaire", "ingenieur", "lecturer", "professor", "administrator"]


def oracle(profil):
    """La logique qu'un bon Claude doit déduire de chaque école."""
    d = {
        "prenom-nom": [("demo-ecole.fr", ["prenom.nom"], ["p.nom"])],
        "numeros": [("essec-demo.edu", ["numero"], ["prenom.nom"])],
        "sous-domaine": [("etu.univ-demo.fr", ["prenom.nom"], []), ("univ-demo.fr", [], ["prenom.nom"])],
        "intitules-seuls": [],
        "nom-prenom": [("ecole-np.fr", ["nom.prenom"], ["prenom.nom"])],
        "multi-domaines": [("campus-lille.demo.fr", ["prenom.nom"], ["p.nom"]), ("campus-paris.demo.fr", ["prenom.nom"], ["p.nom"]),
                           ("ancien-demo.fr", [], ["p.nom"])],
        "anglophone": [("uni-demo.ac.uk", ["prenom.nom"], ["p.nom"])],
        "pieges": [("demo-pieges.fr", ["prenom.nom"], ["p.nom"])],
    }[profil]
    return logique.nettoyer(Logique(domaines=[RegleDomaine(domaine=a, formats_etudiant=e, formats_personnel=p) for a, e, p in d],
                                    mots_etudiant=MOTS_E, mots_personnel=MOTS_P))


class IAPrudente:
    """Le pire cas pour le rappel : l'IA ne tranche jamais (tout « incertain »)."""
    def classer(self, contacts, reperes=None):
        return Lot(resultats={i: ("incertain", "inconnu", "prudence") for i in range(len(contacts))})


# Rappel minimal attendu avec la bonne logique, par école. « intitules-seuls » :
# seuls les étudiants dont l'intitulé est rempli (~70 %) peuvent être reconnus.
RAPPEL_ORACLE = {"prenom-nom": 0.97, "numeros": 0.99, "sous-domaine": 0.99, "intitules-seuls": 0.65,
                 "nom-prenom": 0.97, "multi-domaines": 0.97, "anglophone": 0.97, "pieges": 0.97}


def mesurer(annuaire, resultats):
    verite = {c["adresse"].lower(): c["verite"] for c in annuaire}
    retenus = [r["adresse"] for r in resultats if r["groupe"] == "etudiant"]
    intrus = [(a, verite[a]) for a in retenus if verite[a] != "etudiant"]
    vrais = sum(1 for v in verite.values() if v == "etudiant")
    return intrus, len(retenus) / max(1, vrais)


@pytest.mark.parametrize("profil", list(ecoles.PROFILS))
def test_les_regles_ne_laissent_passer_aucun_piege(profil):
    for c in ecoles.PROFILS[profil]():
        r = regles.appliquer(c["adresse"], c["nom"])
        if c["verite"] in PIEGES:
            assert r is None or r[0] != "etudiant", (profil, c)
        if r and r[0] == "etudiant":
            assert c["verite"] == "etudiant", (profil, c)


@pytest.mark.parametrize("profil", list(ecoles.PROFILS))
def test_pieges_reconnus_par_les_regles(profil):
    """Salles, listes, invités, services, associations et contacts perso : tranchés sans IA."""
    rates = [c for c in ecoles.PROFILS[profil]() if c["verite"] in PIEGES - {"externe"} and regles.appliquer(c["adresse"], c["nom"]) is None]
    assert rates == [], rates[:5]


@pytest.mark.parametrize("profil", list(ecoles.PROFILS))
def test_format_des_adresses_bien_lu(profil):
    """Le format calculé correspond à la convention de l'école, noms composés et accents compris."""
    attendu = {"prenom-nom": ("prenom.nom", "p.nom"), "numeros": ("numero", "prenom.nom"), "sous-domaine": ("prenom.nom", "prenom.nom"),
               "intitules-seuls": ("prenom.nom", "prenom.nom"), "nom-prenom": ("nom.prenom", "prenom.nom"),
               "multi-domaines": ("prenom.nom", "p.nom"), "anglophone": ("prenom.nom", "p.nom"), "pieges": ("prenom.nom", "p.nom")}[profil]
    faux = []
    for c in ecoles.PROFILS[profil]():
        if c["verite"] in ("etudiant", "personnel"):
            f = format_adresse(c["adresse"], c["nom"])
            if f != attendu[0 if c["verite"] == "etudiant" else 1]:
                faux.append((c["adresse"], c["nom"], f))
    assert faux == [], faux[:8]


@pytest.mark.parametrize("profil", list(ecoles.PROFILS))
def test_bonne_logique_zero_intrus_et_bon_rappel(profil, utilisateur):
    service.definir_classeur(IAPrudente())
    annuaire = ecoles.PROFILS[profil]()
    r = service.trier(utilisateur, annuaire, oracle(profil))
    intrus, rappel = mesurer(annuaire, r["resultats"])
    assert intrus == [], intrus[:8]
    assert rappel >= RAPPEL_ORACLE[profil], f"{profil} : {rappel:.1%} des étudiants retrouvés"


@pytest.mark.parametrize("profil", list(ecoles.PROFILS))
def test_logique_apprise_sur_30_puis_ia_prudente_zero_intrus(profil, utilisateur):
    """Comme dans l'app : échantillon de 30 (avec intitulés d'abord), logique apprise, puis tri."""
    service.definir_classeur(IAPrudente())
    annuaire = ecoles.PROFILS[profil]()
    page = annuaire[:200]
    avec = [c for c in page if c["intitule"]]
    sans = [c for c in page if not c["intitule"]]
    echantillon = (avec[:22] + sans[:30 - min(22, len(avec))])[:30]
    lg, _, _ = FauxApprenti().apprendre(echantillon)
    r = service.trier(utilisateur, annuaire, lg)
    intrus, rappel = mesurer(annuaire, r["resultats"])
    assert intrus == [], (profil, lg.model_dump(), intrus[:8])
    print(f"\n{profil:16} logique apprise : {lg.resume}  →  rappel {rappel:.0%}")


def test_le_profil_inverse_ne_trompe_pas():
    """Dans l'école à numéros, prenom.nom = PERSONNEL : la logique d'une autre école ne doit pas fuiter."""
    lg = oracle("prenom-nom")  # logique de demo-ecole.fr
    for c in ecoles.numeros()[:500]:
        r = logique.appliquer(c, lg)
        if r and r[0] == "etudiant":
            # seuls les intitulés étudiants peuvent trancher hors du domaine appris
            assert c["verite"] == "etudiant", c
