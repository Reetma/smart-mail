"""Outils communs aux tests : base SQLite neuve par test, faux client Anthropic
qui imite la vraie API (blocs de contenu, usage, stop_reason) et qui classe les
contacts avec une logique d'« école » réaliste."""

import json
import re
import threading
import time
from types import SimpleNamespace

import httpx2
import anthropic
import pytest

from app import base, config, service


@pytest.fixture(autouse=True)
def base_neuve(tmp_path, monkeypatch):
    base.fermer()
    monkeypatch.setattr(config, "BASE", tmp_path / "test.db")
    monkeypatch.setattr(config, "BUDGET_MENSUEL_USD", 50.0)
    monkeypatch.setattr(config, "TAILLE_LOT", 150)
    monkeypatch.setattr(config, "PARALLELE", 4)
    service.definir_classeur(None)
    yield
    service.definir_classeur(None)
    base.fermer()


def erreur_api(classe, statut):
    req = httpx2.Request("POST", "https://api.anthropic.com/v1/messages")
    if classe is anthropic.APIConnectionError:
        return anthropic.APIConnectionError(request=req)
    return classe("erreur simulée", response=httpx2.Response(statut, request=req), body=None)


def lire_lot(message):
    """Retrouve (i, adresse, nom, intitulé) dans le message envoyé à Claude."""
    sortie = []
    for l in message.splitlines():
        parts = l.split(" | ")
        if len(parts) == 4 and parts[0].isdigit():
            sortie.append((int(parts[0]), parts[1].strip(), parts[2].strip(), parts[3].strip()))
    return sortie


def logique_ecole(i, adresse, nom, intitule):
    """Ce qu'on attend d'un bon tri : l'intitulé d'abord, puis le format d'adresse."""
    t = (nom + " " + intitule).lower()
    locale = adresse.split("@")[0]
    if any(m in t for m in ("professeur", "prof.", "enseignant", "dr ", "directrice", "directeur", "chargée de", "chargé de")):
        return {"i": i, "groupe": "hors", "type": "personnel", "raison": "fonction du personnel"}
    if any(m in t for m in ("étudiant", "etudiant", "programme", "promo", "bachelor")):
        return {"i": i, "groupe": "etudiant", "type": "etudiant", "raison": "intitulé étudiant"}
    if re.match(r"^[a-z]\.[a-z-]+$", locale):
        return {"i": i, "groupe": "hors", "type": "personnel", "raison": "format initiale.nom du personnel"}
    if re.match(r"^[a-z-]{2,}\.[a-z-]+$", locale):
        return {"i": i, "groupe": "etudiant", "type": "etudiant", "raison": "format prenom.nom des étudiants"}
    if nom.isupper() and len(nom) <= 12:
        return {"i": i, "groupe": "hors", "type": "association", "raison": "sigle, pas une personne"}
    return {"i": i, "groupe": "incertain", "type": "inconnu", "raison": "aucun indice"}


class FauxAnthropic:
    """Imite anthropic.Anthropic() : client.beta.messages.create(...)."""

    def __init__(self, comportement=None, latence=0.0):
        self.appels = []
        self.comportement = comportement  # fonction(appel_n, contacts) -> réponse ou exception
        self.latence = latence
        self._verrou = threading.Lock()
        self.beta = SimpleNamespace(messages=SimpleNamespace(create=self._create))

    def _create(self, **kw):
        with self._verrou:
            self.appels.append(kw)
            n = len(self.appels)
        if self.latence:
            time.sleep(self.latence)
        contacts = lire_lot(kw["messages"][0]["content"])
        if self.comportement:
            r = self.comportement(n, contacts, kw)
            if isinstance(r, Exception):
                raise r
            if r is not None:
                return r
        return reponse([logique_ecole(*c) for c in contacts])


def reponse(resultats=None, stop_reason="end_turn", texte=None, entree=1200, sortie=600, modele="claude-opus-5-5"):
    t = texte if texte is not None else json.dumps({"resultats": resultats or []}, ensure_ascii=False)
    return SimpleNamespace(
        content=[SimpleNamespace(type="thinking", thinking=""), SimpleNamespace(type="text", text=t)],
        stop_reason=stop_reason,
        model=modele,
        usage=SimpleNamespace(input_tokens=entree, output_tokens=sortie, cache_read_input_tokens=0, cache_creation_input_tokens=0),
    )


@pytest.fixture
def utilisateur():
    jeton = base.creer_utilisateur("Test", 20000)
    return base.utilisateur_par_jeton(jeton)
