"""API du serveur Smart Mail.

  GET  /sante      état du serveur (sans authentification)
  GET  /v1/moi     qui suis-je, quota et consommation du mois
  POST /v1/tri     trie une liste de contacts : {"contacts": [{"adresse", "nom", "intitule"}]}

Authentification : en-tête « Authorization: Bearer <code d'accès> » (codes créés
avec `python -m app.admin creer`).
"""

import os
from pathlib import Path

from fastapi import Depends, FastAPI, Header, HTTPException
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from . import base, config, demo, logique as logique_mod, service

app = FastAPI(title="Smart Mail", version="1.0.0", docs_url=None, redoc_url=None)
app.add_middleware(
    CORSMiddleware,
    allow_origins=config.ORIGINES,
    allow_methods=["GET", "POST"],
    allow_headers=["Authorization", "Content-Type"],
)


class Contact(BaseModel):
    adresse: str = Field(max_length=320)
    nom: str = Field(default="", max_length=300)
    intitule: str = Field(default="", max_length=300)


class DemandeTri(BaseModel):
    contacts: list[Contact]
    logique: logique_mod.Logique | None = None


class Destinataire(BaseModel):
    adresse: str = Field(max_length=320)
    nom: str = Field(default="", max_length=300)
    source: str = Field(default="", max_length=20)


class AjoutDestinataires(BaseModel):
    personnes: list[Destinataire] = Field(max_length=2000)


class EtatRecherche(BaseModel):
    """Où en est la recherche : de quoi la reprendre sur n'importe quel appareil."""
    objectif: int = Field(ge=1, le=100000)
    depart: int | None = Field(default=0, ge=0)
    fini: bool = False
    ecartes: int = Field(default=0, ge=0)
    logique: logique_mod.Logique | None = None
    resume: str = Field(default="", max_length=400)


class Parametres(BaseModel):
    """Paramètres d'envoi, gardés dans le compte."""
    taille_pack: int = Field(default=1, ge=1, le=100)       # destinataires par message (1 = individuel, sinon Cci)
    packs_heure: int = Field(default=35, ge=1, le=120)      # messages envoyés par heure
    max_jour: int = Field(default=400, ge=1, le=10000)      # destinataires au plus sur 24 h
    heure_debut: int = Field(default=8, ge=0, le=23)
    heure_fin: int = Field(default=20, ge=1, le=24)
    preset: str = Field(default="standard", max_length=20)


class DemandeLogique(BaseModel):
    echantillon: list[Contact] = Field(min_length=5, max_length=80)


_apprenti = None


def apprenti():
    global _apprenti
    if _apprenti is None:
        _apprenti = logique_mod.apprenti()
    return _apprenti


def definir_apprenti(a):
    """Pour les tests."""
    global _apprenti
    _apprenti = a


def utilisateur(authorization: str = Header(default="")):
    jeton = authorization[7:].strip() if authorization.lower().startswith("bearer ") else ""
    u = base.utilisateur_par_jeton(jeton)
    if not u:
        raise HTTPException(status_code=401, detail="Code d'accès invalide ou désactivé.")
    return u


@app.get("/sante")
def sante():
    return {
        "ok": True,
        "ia": "simulation" if config.FAUX_CLAUDE else ("claude" if os.environ.get("ANTHROPIC_API_KEY") else "clé manquante"),
        "modele": config.MODELE,
    }


@app.get("/v1/moi")
def moi(u=Depends(utilisateur)):
    g = base.usage(u["id"])
    return {"nom": u["nom"], "quota": {"mensuel": u["quota_mensuel"], "utilise": g["contacts_ia"]},
            "contacts_tries_ce_mois": g["contacts_total"]}


@app.get("/v1/compte")
def compte(u=Depends(utilisateur)):
    """Tout ce qu'il faut pour reprendre là où on en était, sur n'importe quel appareil."""
    return {"nom": u["nom"], "email": u.get("email") or "", "destinataires": base.destinataires_compter(u["id"]),
            "recherche": base.recherche_lire(u["id"]), "parametres": base.parametres_lire(u["id"]) or Parametres().model_dump()}


@app.get("/v1/destinataires")
def lister_destinataires(depart: int = 0, taille: int = 500, u=Depends(utilisateur)):
    taille = max(1, min(taille, 2000))
    return {"personnes": base.destinataires_lister(u["id"], max(0, depart), taille), "total": base.destinataires_compter(u["id"])}


@app.post("/v1/destinataires")
def ajouter_destinataires(demande: AjoutDestinataires, u=Depends(utilisateur)):
    return {"total": base.destinataires_ajouter(u["id"], [p.model_dump() for p in demande.personnes])}


@app.delete("/v1/destinataires")
def vider_destinataires(u=Depends(utilisateur)):
    base.destinataires_vider(u["id"])
    return {"total": 0}


@app.put("/v1/recherche")
def enregistrer_recherche(etat: EtatRecherche, u=Depends(utilisateur)):
    base.recherche_ecrire(u["id"], etat.model_dump())
    return {"ok": True}


@app.put("/v1/parametres")
def enregistrer_parametres(p: Parametres, u=Depends(utilisateur)):
    if p.heure_fin <= p.heure_debut:
        raise HTTPException(status_code=422, detail="L'heure de fin doit suivre l'heure de début.")
    base.parametres_ecrire(u["id"], p.model_dump())
    return p.model_dump()


@app.post("/v1/deconnexion")
def deconnexion(authorization: str = Header(default=""), u=Depends(utilisateur)):
    base.deconnecter(authorization[7:].strip())
    return {"ok": True}


@app.post("/v1/tri")
async def tri(demande: DemandeTri, u=Depends(utilisateur)):
    if len(demande.contacts) > config.MAX_CONTACTS_PAR_REQUETE:
        raise HTTPException(status_code=413, detail=f"{config.MAX_CONTACTS_PAR_REQUETE} contacts au plus par envoi.")
    if not config.FAUX_CLAUDE and not os.environ.get("ANTHROPIC_API_KEY"):
        raise HTTPException(status_code=503, detail="Le tri IA n'est pas encore configuré sur le serveur (clé manquante).")
    # Les appels à Claude sont bloquants : hors de la boucle asynchrone.
    lg = logique_mod.nettoyer(demande.logique) if demande.logique else None
    return await run_in_threadpool(service.trier, u, [c.model_dump() for c in demande.contacts], lg)


@app.post("/v1/logique")
async def apprendre_logique(demande: DemandeLogique, u=Depends(utilisateur)):
    """Étape de lancement : Claude déduit la logique de l'école d'un échantillon (~30 personnes)."""
    if not config.FAUX_CLAUDE and not os.environ.get("ANTHROPIC_API_KEY"):
        raise HTTPException(status_code=503, detail="Le tri IA n'est pas encore configuré sur le serveur (clé manquante).")
    if base.depense_du_mois() >= config.BUDGET_MENSUEL_USD:
        raise HTTPException(status_code=429, detail="Budget IA du mois atteint.")
    lg, cout, erreur = await run_in_threadpool(apprenti().apprendre, [c.model_dump() for c in demande.echantillon])
    base.ajouter_usage(u["id"], cout_usd=cout)
    if lg is None:
        raise HTTPException(status_code=502, detail=erreur or "logique introuvable")
    return {"logique": lg.model_dump(), "resume": lg.resume, "cout_usd": round(cout, 4)}


@app.get("/v1/config")
def config_publique():
    """Ce que l'application a besoin de savoir au démarrage."""
    return {"demo": config.DEMO}


app.include_router(demo.router)
# L'application (PWA) : servie à la racine, après les routes de l'API.
app.mount("/", StaticFiles(directory=Path(__file__).parent / "web", html=True), name="web")
