import json

import pytest
from fastapi.testclient import TestClient

from app import base, config, logique, service
from app.ia import ClasseurClaude
from app.logique import ApprentiClaude, FauxApprenti, Logique, RegleDomaine, appliquer, format_adresse, nettoyer
from app.main import app, definir_apprenti
from conftest import FauxAnthropic, reponse


@pytest.mark.parametrize("adresse,nom,attendu", [
    ("marie.durand@ieseg.fr", "Marie Durand", "prenom.nom"),
    ("marie.durand2@ieseg.fr", "Marie Durand", "prenom.nom"),
    ("anne-laure.gautier@ieseg.fr", "Anne-Laure Gautier", "prenom.nom"),
    ("durand.marie@ieseg.fr", "Marie Durand", "nom.prenom"),
    ("m.durand@ieseg.fr", "Marie Durand", "p.nom"),
    ("mdurand@ieseg.fr", "Marie Durand", "pnom"),
    ("mariedurand@ieseg.fr", "Marie Durand", "prenomnom"),
    ("marie@ieseg.fr", "Marie Durand", "prenom"),
    ("durand@ieseg.fr", "Marie Durand", "nom"),
    ("helene.lefevre@ieseg.fr", "Hélène Lefèvre", "prenom.nom"),
    ("j.roux@ieseg.fr", "Prof. Jean Roux", "p.nom"),
    ("b00654321@essec.edu", "", "numero"),
    ("m.durand@ieseg.fr", "", "p.nom"),
    ("lucas.petit@ieseg.fr", "", "prenom.nom"),
    ("xyz123abc@ieseg.fr", "", "autre"),
])
def test_format_adresse(adresse, nom, attendu):
    assert format_adresse(adresse, nom) == attendu


def logique_ieseg():
    return nettoyer(Logique(
        domaines=[RegleDomaine(domaine="IESEG.fr", formats_etudiant=["prenom.nom", "bizarre"], formats_personnel=["p.nom", "autre"])],
        mots_etudiant=["Étudiant", "PGE"], mots_personnel=["Professeur", "ab"], resume="x"))


def test_nettoyer():
    lg = logique_ieseg()
    assert lg.domaines[0].domaine == "ieseg.fr"
    assert lg.domaines[0].formats_etudiant == ["prenom.nom"] and lg.domaines[0].formats_personnel == ["p.nom"]
    assert lg.mots_etudiant == ["etudiant", "pge"] and lg.mots_personnel == ["professeur"]
    conflit = nettoyer(Logique(domaines=[RegleDomaine(domaine="a.fr", formats_etudiant=["prenom.nom"], formats_personnel=["prenom.nom", "p.nom"])]))
    assert conflit.domaines[0].formats_etudiant == [] and conflit.domaines[0].formats_personnel == ["p.nom"]


def test_appliquer():
    lg = logique_ieseg()
    assert appliquer({"adresse": "marie.durand@ieseg.fr", "nom": "Marie Durand"}, lg)[0] == "etudiant"
    assert appliquer({"adresse": "m.durand@ieseg.fr", "nom": "Marie Durand"}, lg)[:2] == ("hors", "personnel")
    # L'intitulé l'emporte sur le format
    assert appliquer({"adresse": "marie.durand@ieseg.fr", "nom": "Marie Durand", "intitule": "Professeur associé"}, lg)[1] == "personnel"
    assert appliquer({"adresse": "m.durand@ieseg.fr", "intitule": "Étudiante PGE 2027"}, lg)[0] == "etudiant"
    # « pge » doit être un mot entier
    assert appliquer({"adresse": "zz@autre.fr", "intitule": "Pgeneral"}, lg) is None
    # Format non couvert, autre domaine, pas de logique : on laisse à Claude
    assert appliquer({"adresse": "mdurand@ieseg.fr", "nom": "Marie Durand"}, lg) is None
    assert appliquer({"adresse": "marie.durand@essca.fr", "nom": "Marie Durand"}, lg) is None
    assert appliquer({"adresse": "marie.durand@ieseg.fr"}, None) is None


def echantillon_ecole():
    e = [{"adresse": f"{p}.{n}@ieseg.fr", "nom": f"{p.title()} {n.title()}", "intitule": "Étudiant PGE"}
         for p, n in [("marie", "durand"), ("lucas", "petit"), ("ines", "roux"), ("hugo", "morel")]]
    p = [{"adresse": f"{p[0]}.{n}@ieseg.fr", "nom": f"{p.title()} {n.title()}", "intitule": "Professeur de finance"}
         for p, n in [("jean", "martin"), ("lina", "benali"), ("paul", "girard")]]
    autres = [{"adresse": f"anon{i}@ieseg.fr", "nom": "", "intitule": ""} for i in range(3)]
    return e + p + autres


def test_faux_apprenti_trouve_la_logique():
    lg, cout, err = FauxApprenti().apprendre(echantillon_ecole())
    d = lg.domaines[0]
    assert d.domaine == "ieseg.fr" and d.formats_etudiant == ["prenom.nom"] and d.formats_personnel == ["p.nom"]
    assert "prenom.nom@ieseg.fr" in lg.resume and err == ""


def test_apprenti_claude_requete_et_lecture():
    rendu = {"domaines": [{"domaine": "ieseg.fr", "formats_etudiant": ["prenom.nom"], "formats_personnel": ["p.nom"]}],
             "mots_etudiant": ["étudiant", "PGE"], "mots_personnel": ["professeur"],
             "resume": "Étudiants en prenom.nom@ieseg.fr, personnel en p.nom@ieseg.fr."}
    faux = FauxAnthropic(lambda n, c, k: reponse(texte=json.dumps(rendu)))
    lg, cout, err = ApprentiClaude(client=faux).apprendre(echantillon_ecole())
    kw = faux.appels[0]
    assert kw["model"] == "claude-opus-5-5" and kw["output_config"]["effort"] == "medium"
    assert kw["output_config"]["format"]["schema"] == logique.SCHEMA and kw["fallbacks"] == "default"
    msg = kw["messages"][0]["content"]
    assert "marie.durand@ieseg.fr | prenom.nom | Marie Durand | Étudiant PGE" in msg
    assert "j.martin@ieseg.fr | p.nom |" in msg
    assert err == "" and cout > 0 and lg.mots_etudiant == ["etudiant", "pge"]


def test_apprenti_claude_echecs():
    for comp, texte in [(lambda n, c, k: reponse(texte="pas du json"), "illisible"),
                        (lambda n, c, k: reponse(stop_reason="refusal", texte=""), "pas pu"),
                        (lambda n, c, k: reponse(texte='{"domaines": "x"}'), "illisible")]:
        lg, cout, err = ApprentiClaude(client=FauxAnthropic(comp)).apprendre(echantillon_ecole())
        assert lg is None and texte in err


def test_la_logique_evite_les_appels_a_claude(utilisateur):
    faux = FauxAnthropic()
    service.definir_classeur(ClasseurClaude(client=faux))
    lg, _, _ = FauxApprenti().apprendre(echantillon_ecole())
    contacts = [{"adresse": f"prenom{x}.nom{x}@ieseg.fr", "nom": f"Prenom{x} Nom{x}"} for x in "abcdefghij"] + \
               [{"adresse": f"{x}.nom{x}@ieseg.fr", "nom": f"{x.upper()}arc Nom{x}"} for x in "klmno"] + \
               [{"adresse": "tom@ieseg.fr", "nom": "Tom"}]
    r = service.trier(utilisateur, contacts, lg)
    src = [x["source"] for x in r["resultats"]]
    assert src.count("logique") == 15 and src.count("ia") == 1
    assert len(faux.appels) == 1 and faux.appels[0]["messages"][0]["content"].count("@ieseg.fr |") == 1
    assert r["resume"] == {"etudiant": 10, "hors": 5, "incertain": 1}
    assert base.cache_lire(["prenoma.noma@ieseg.fr"]) == {}  # pas de cache pour la logique


def test_api_logique_puis_tri(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-test-factice")
    definir_apprenti(FauxApprenti())
    service.definir_classeur(ClasseurClaude(client=FauxAnthropic()))
    code = base.creer_utilisateur("Camille", 1000)
    c, h = TestClient(app), {"Authorization": f"Bearer {code}"}
    assert c.post("/v1/logique", headers=h, json={"echantillon": echantillon_ecole()[:3]}).status_code == 422  # trop petit
    r = c.post("/v1/logique", headers=h, json={"echantillon": echantillon_ecole()})
    assert r.status_code == 200 and "prenom.nom@ieseg.fr" in r.json()["resume"]
    lg = r.json()["logique"]
    r2 = c.post("/v1/tri", headers=h, json={"contacts": [{"adresse": "zoe.lambert@ieseg.fr", "nom": "Zoé Lambert"},
                                                          {"adresse": "z.lambert@ieseg.fr", "nom": "Zoé Lambert"}], "logique": lg})
    assert [(x["groupe"], x["source"]) for x in r2.json()["resultats"]] == [("etudiant", "logique"), ("hors", "logique")]
    # Logique malformée : refusée proprement
    assert c.post("/v1/tri", headers=h, json={"contacts": [], "logique": {"domaines": "x"}}).status_code == 422
    definir_apprenti(None)


def test_api_logique_budget_atteint(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-test-factice")
    monkeypatch.setattr(config, "BUDGET_MENSUEL_USD", 0.0)
    definir_apprenti(FauxApprenti())
    code = base.creer_utilisateur("Camille", 1000)
    r = TestClient(app).post("/v1/logique", headers={"Authorization": f"Bearer {code}"}, json={"echantillon": echantillon_ecole()})
    assert r.status_code == 429
    definir_apprenti(None)
