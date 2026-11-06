"""Mode démonstration (SMARTMAIL_DEMO=1) : un faux annuaire d'école réaliste,
servi comme le serait l'annuaire Microsoft (Graph /users), et un accès de
démonstration. Permet de montrer et tester toute l'application sans compte
Microsoft ni clé Anthropic. Aucun mail n'est jamais envoyé."""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

from . import base, config, ecoles

router = APIRouter(prefix="/demo")

_ANNUAIRES = {}


def _verifier():
    if not config.DEMO:
        raise HTTPException(status_code=404, detail="Mode démonstration désactivé.")


@router.get("/annuaire")
def page(depart: int = 0, taille: int = 200, ecole: str = "prenom-nom"):
    """Une page de l'annuaire d'une école fictive (app/ecoles.py), comme le renverrait
    Microsoft Graph. Le champ « verite » sert aux tests ; le tri l'ignore."""
    _verifier()
    if ecole not in ecoles.PROFILS:
        raise HTTPException(status_code=404, detail="École de démonstration inconnue.")
    if ecole not in _ANNUAIRES:
        _ANNUAIRES[ecole] = ecoles.PROFILS[ecole]()
    a = _ANNUAIRES[ecole]
    depart, taille = max(0, depart), max(1, min(taille, 500))
    return {"personnes": a[depart:depart + taille], "suivant": depart + taille if depart + taille < len(a) else None, "total": len(a)}


@router.get("/ecoles")
def liste_ecoles():
    _verifier()
    return {"ecoles": list(ecoles.PROFILS)}


class DemandeSession(BaseModel):
    email: str = Field(min_length=3, max_length=200, pattern=r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


@router.post("/session")
def session(demande: DemandeSession):
    """Connexion de démonstration par adresse e-mail : même adresse, même compte
    (sur n'importe quel appareil). En production, c'est la connexion Microsoft qui
    crée et retrouve le compte."""
    _verifier()
    jeton, u = base.connexion_par_email(demande.email, quota=100000)
    if not jeton:
        raise HTTPException(status_code=403, detail="Compte désactivé.")
    return {"code": jeton, "compte": u["email"], "nom": u["nom"]}
