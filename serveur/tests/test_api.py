import pytest
from fastapi.testclient import TestClient

from app import admin, base, config, service
from app.ia import ClasseurClaude, FauxClasseur
from app.main import app
from conftest import FauxAnthropic


@pytest.fixture
def client(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-ant-test-factice")
    service.definir_classeur(ClasseurClaude(client=FauxAnthropic()))
    return TestClient(app)


@pytest.fixture
def code():
    return base.creer_utilisateur("Camille", 1000)


def test_sante_sans_authentification(client):
    r = client.get("/sante")
    assert r.status_code == 200 and r.json()["ok"] is True and r.json()["ia"] == "claude"
    assert "sk-ant" not in r.text  # la clé n'est jamais renvoyée


def test_codes_refuses(client, code):
    for h in ({}, {"Authorization": "Bearer faux"}, {"Authorization": code}, {"Authorization": "Basic " + code}):
        assert client.post("/v1/tri", json={"contacts": []}, headers=h).status_code == 401
    base.modifier("Camille", actif=0)
    assert client.get("/v1/moi", headers={"Authorization": f"Bearer {code}"}).status_code == 401


def test_tri_par_l_api(client, code):
    h = {"Authorization": f"Bearer {code}"}
    r = client.post("/v1/tri", headers=h, json={"contacts": [
        {"adresse": "marie.durand@ieseg.fr", "nom": "Marie Durand"},
        {"adresse": "bde@ieseg.fr"},
        {"adresse": "x@gmail.com", "intitule": ""},
    ]})
    assert r.status_code == 200, r.text
    j = r.json()
    assert [x["groupe"] for x in j["resultats"]] == ["etudiant", "hors", "hors"]
    assert j["resume"] == {"etudiant": 1, "hors": 2, "incertain": 0}
    moi = client.get("/v1/moi", headers=h).json()
    assert moi == {"nom": "Camille", "quota": {"mensuel": 1000, "utilise": 1}, "contacts_tries_ce_mois": 3}


def test_trop_de_contacts(client, code, monkeypatch):
    monkeypatch.setattr(config, "MAX_CONTACTS_PAR_REQUETE", 3)
    r = client.post("/v1/tri", headers={"Authorization": f"Bearer {code}"}, json={"contacts": [{"adresse": f"a{i}@b.fr"} for i in range(4)]})
    assert r.status_code == 413


def test_champs_trop_longs_refuses(client, code):
    r = client.post("/v1/tri", headers={"Authorization": f"Bearer {code}"}, json={"contacts": [{"adresse": "a" * 400 + "@b.fr"}]})
    assert r.status_code == 422


def test_sans_cle_le_tri_est_desactive_proprement(code, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.setattr(config, "FAUX_CLAUDE", False)
    c = TestClient(app)
    assert c.get("/sante").json()["ia"] == "clé manquante"
    r = c.post("/v1/tri", headers={"Authorization": f"Bearer {code}"}, json={"contacts": [{"adresse": "a.b@c.fr"}]})
    assert r.status_code == 503 and "clé manquante" in r.json()["detail"]


def test_mode_simulation(code, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    monkeypatch.setattr(config, "FAUX_CLAUDE", True)
    service.definir_classeur(FauxClasseur())
    c = TestClient(app)
    assert c.get("/sante").json()["ia"] == "simulation"
    r = c.post("/v1/tri", headers={"Authorization": f"Bearer {code}"}, json={"contacts": [
        {"adresse": "b00654321@essec.edu"}, {"adresse": "lea.petit@essec.edu"},
        {"adresse": "j.roux@essec.edu", "nom": "Prof. Jean Roux"}, {"adresse": "s.morel@essec.edu", "intitule": "Étudiante PGE"}]})
    # « lea.petit » : un repère numérique ne dit rien du format prénom.nom → prudence, incertain.
    assert [x["groupe"] for x in r.json()["resultats"]] == ["etudiant", "incertain", "hors", "etudiant"]


def test_cors_outlook(client):
    r = client.options("/v1/tri", headers={"Origin": "https://outlook.office.com", "Access-Control-Request-Method": "POST",
                                           "Access-Control-Request-Headers": "authorization,content-type"})
    assert r.headers.get("access-control-allow-origin") == "https://outlook.office.com"
    r = client.options("/v1/tri", headers={"Origin": "https://pirate.example", "Access-Control-Request-Method": "POST"})
    assert r.headers.get("access-control-allow-origin") is None


def test_admin_cli(capsys):
    assert admin.main(["creer", "Pierre", "--quota", "500"]) == 0
    sortie = capsys.readouterr().out
    code = sortie.strip().splitlines()[-1]
    assert code.startswith("sm_") and base.utilisateur_par_jeton(code)["quota_mensuel"] == 500
    assert admin.main(["creer", "Pierre"]) == 1  # déjà existant
    assert admin.main(["quota", "Pierre", "900"]) == 0
    assert admin.main(["code", "Pierre"]) == 0
    nouveau = capsys.readouterr().out.strip().splitlines()[-1]
    assert base.utilisateur_par_jeton(code) is None and base.utilisateur_par_jeton(nouveau)["quota_mensuel"] == 900
    assert admin.main(["desactiver", "Pierre"]) == 0 and base.utilisateur_par_jeton(nouveau) is None
    assert admin.main(["activer", "Pierre"]) == 0 and base.utilisateur_par_jeton(nouveau)
    assert admin.main(["quota", "Personne", "1"]) == 1
    assert admin.main(["liste"]) == 0
    assert "Pierre" in capsys.readouterr().out
    assert "sm_" not in str(base.utilisateurs())  # aucun code en clair dans la base


def test_demo_desactivee_par_defaut(client, monkeypatch):
    monkeypatch.setattr(config, "DEMO", False)
    assert client.get("/demo/annuaire").status_code == 404
    assert client.post("/demo/session", json={"email": "a@b.fr"}).status_code == 404
    assert client.get("/v1/config").json() == {"demo": False}


def test_demo_annuaire_et_session(client, monkeypatch):
    monkeypatch.setattr(config, "DEMO", True)
    p = client.get("/demo/annuaire?depart=3990&taille=50").json()
    assert p["total"] == 4000 and len(p["personnes"]) == 10 and p["suivant"] is None
    assert client.get("/demo/annuaire?ecole=numeros").json()["total"] == 3000
    assert client.get("/demo/annuaire?ecole=inconnue").status_code == 404
    assert len(client.get("/demo/ecoles").json()["ecoles"]) == 8
    assert client.get("/demo/annuaire?taille=9999").json()["suivant"] == 500  # taille plafonnée
    s = client.post("/demo/session", json={"email": "Camille@Ecole.fr"}).json()
    assert s["compte"] == "camille@ecole.fr"
    assert s["code"].startswith("sm_") and client.get("/v1/moi", headers={"Authorization": "Bearer " + s["code"]}).status_code == 200


def test_application_servie(client):
    r = client.get("/")
    assert r.status_code == 200 and "Smart Mail" in r.text and 'rel="manifest"' in r.text
    assert client.get("/manifest.webmanifest").json()["display"] == "standalone"
    assert client.get("/sw.js").status_code == 200 and client.get("/icones/icone-512.png").status_code == 200



def test_compte_le_meme_sur_deux_appareils_et_memoire(client, monkeypatch):
    monkeypatch.setattr(config, "DEMO", True)
    a = client.post("/demo/session", json={"email": "lea@ecole.fr"}).json()["code"]
    b = client.post("/demo/session", json={"email": "LEA@ecole.fr"}).json()["code"]  # 2e appareil
    assert a != b
    ha, hb = {"Authorization": "Bearer " + a}, {"Authorization": "Bearer " + b}
    c0 = client.get("/v1/compte", headers=ha).json()
    assert c0["email"] == "lea@ecole.fr" and c0["destinataires"] == 0 and c0["recherche"] is None
    assert c0["parametres"]["taille_pack"] == 1 and c0["parametres"]["max_jour"] == 400

    # Le téléphone trouve des étudiants…
    r = client.post("/v1/destinataires", headers=ha, json={"personnes": [
        {"adresse": "Marie.Durand@ecole.fr", "nom": "Marie Durand", "source": "logique"},
        {"adresse": "lucas.petit@ecole.fr", "nom": "Lucas Petit", "source": "ia"}]})
    assert r.json()["total"] == 2
    assert client.post("/v1/destinataires", headers=ha, json={"personnes": [{"adresse": "marie.durand@ecole.fr"}]}).json()["total"] == 2
    lg = {"domaines": [{"domaine": "ecole.fr", "formats_etudiant": ["prenom.nom"], "formats_personnel": []}], "resume": "x"}
    assert client.put("/v1/recherche", headers=ha, json={"objectif": 500, "depart": 400, "ecartes": 37, "logique": lg}).status_code == 200

    # …et l'ordinateur retrouve tout.
    c = client.get("/v1/compte", headers=hb).json()
    assert c["destinataires"] == 2 and c["recherche"]["objectif"] == 500 and c["recherche"]["depart"] == 400
    assert c["recherche"]["logique"]["domaines"][0]["formats_etudiant"] == ["prenom.nom"]
    l = client.get("/v1/destinataires", headers=hb).json()
    assert [p["adresse"] for p in l["personnes"]] == ["marie.durand@ecole.fr", "lucas.petit@ecole.fr"]

    # Les listes sont privées : un autre compte ne voit rien.
    autre = {"Authorization": "Bearer " + client.post("/demo/session", json={"email": "tom@ecole.fr"}).json()["code"]}
    assert client.get("/v1/destinataires", headers=autre).json()["total"] == 0

    # Nouvelle recherche : tout est vidé.
    client.delete("/v1/destinataires", headers=hb)
    assert client.get("/v1/compte", headers=ha).json()["destinataires"] == 0
    assert client.get("/v1/compte", headers=ha).json()["recherche"] is None

    # Déconnexion d'un appareil : l'autre reste connecté.
    assert client.post("/v1/deconnexion", headers=hb).status_code == 200
    assert client.get("/v1/compte", headers=hb).status_code == 401
    assert client.get("/v1/compte", headers=ha).status_code == 200


def test_parametres_d_envoi(client, monkeypatch):
    monkeypatch.setattr(config, "DEMO", True)
    h = {"Authorization": "Bearer " + client.post("/demo/session", json={"email": "p@ecole.fr"}).json()["code"]}
    p = {"taille_pack": 30, "packs_heure": 4, "max_jour": 900, "heure_debut": 9, "heure_fin": 18, "preset": "perso"}
    assert client.put("/v1/parametres", headers=h, json=p).json() == p
    assert client.get("/v1/compte", headers=h).json()["parametres"] == p
    for mauvais in ({"taille_pack": 0}, {"taille_pack": 101}, {"packs_heure": 500}, {"max_jour": 0}, {"heure_debut": 18, "heure_fin": 9}):
        assert client.put("/v1/parametres", headers=h, json={**p, **mauvais}).status_code == 422, mauvais
    assert client.get("/v1/compte", headers=h).json()["parametres"] == p  # inchangé


def test_session_email_invalide(client, monkeypatch):
    monkeypatch.setattr(config, "DEMO", True)
    for e in ("", "pas-un-email", "a b@c.fr"):
        assert client.post("/demo/session", json={"email": e}).status_code == 422


def test_migration_base_existante(tmp_path, monkeypatch):
    """Une base créée par la première version (sans email ni parametres) est mise à niveau."""
    import sqlite3
    chemin = tmp_path / "ancienne.db"
    c = sqlite3.connect(chemin)
    c.executescript("CREATE TABLE utilisateurs (id INTEGER PRIMARY KEY, nom TEXT NOT NULL UNIQUE, jeton_hash TEXT NOT NULL UNIQUE, "
                    "quota_mensuel INTEGER NOT NULL, actif INTEGER NOT NULL DEFAULT 1, cree TEXT NOT NULL);"
                    "INSERT INTO utilisateurs (nom, jeton_hash, quota_mensuel, cree) VALUES ('Ancien', 'x', 10, '2026-01-01');")
    c.commit(); c.close()
    base.fermer()
    monkeypatch.setattr(config, "BASE", chemin)
    assert [u["nom"] for u in base.utilisateurs()] == ["Ancien"]
    jeton, u = base.connexion_par_email("nouveau@ecole.fr")
    assert base.utilisateur_par_jeton(jeton)["email"] == "nouveau@ecole.fr"
