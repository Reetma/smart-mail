from app import base, config, service
from app.ia import ClasseurClaude
from conftest import FauxAnthropic, erreur_api, reponse

import anthropic


def brancher(faux):
    service.definir_classeur(ClasseurClaude(client=faux))
    return faux


def contacts_ecole():
    return [
        {"adresse": "Marie.Durand@ieseg.fr", "nom": "Marie Durand"},
        {"adresse": "p.martin@ieseg.fr", "nom": "Pierre Martin"},
        {"adresse": "lbenali@ieseg.fr", "nom": "Lina Benali", "intitule": "Professeure de finance"},
        {"adresse": "bde@ieseg.fr", "nom": "BDE IESEG"},
        {"adresse": "baptiste.rousseau@gmail.com", "nom": "Baptiste Rousseau"},
        {"adresse": "b00654321@essec.edu", "nom": "Léa Petit"},
        {"adresse": "tom@ieseg.fr", "nom": "Tom"},
        {"adresse": "marie.durand@ieseg.fr", "nom": ""},  # doublon (casse différente)
    ]


def test_tri_complet_regles_puis_claude(utilisateur):
    faux = brancher(FauxAnthropic())
    r = service.trier(utilisateur, contacts_ecole())
    par = {x["adresse"]: x for x in r["resultats"]}
    assert list(par) == ["marie.durand@ieseg.fr", "p.martin@ieseg.fr", "lbenali@ieseg.fr", "bde@ieseg.fr",
                         "baptiste.rousseau@gmail.com", "b00654321@essec.edu", "tom@ieseg.fr"]  # ordre gardé, doublon fusionné
    assert par["bde@ieseg.fr"]["source"] == "regle" and par["bde@ieseg.fr"]["type"] == "association"
    assert par["baptiste.rousseau@gmail.com"]["type"] == "perso"
    assert par["b00654321@essec.edu"]["groupe"] == "etudiant" and par["b00654321@essec.edu"]["source"] == "regle"
    assert par["marie.durand@ieseg.fr"]["groupe"] == "etudiant" and par["marie.durand@ieseg.fr"]["source"] == "ia"
    assert par["p.martin@ieseg.fr"]["type"] == "personnel"
    assert par["lbenali@ieseg.fr"]["type"] == "personnel"
    assert par["tom@ieseg.fr"]["groupe"] == "incertain"
    assert r["resume"] == {"etudiant": 2, "hors": 4, "incertain": 1}
    assert len(faux.appels) == 1  # 4 contacts → un seul lot pour Claude
    msg = faux.appels[0]["messages"][0]["content"]
    assert "bde@ieseg.fr" not in msg and "gmail.com" not in msg  # les règles ne coûtent rien
    assert r["quota"]["utilise"] == 4 and r["cout_usd"] > 0


def test_le_cache_evite_de_repayer(utilisateur):
    faux = brancher(FauxAnthropic())
    service.trier(utilisateur, contacts_ecole())
    r2 = service.trier(utilisateur, contacts_ecole() + [{"adresse": "julie.roux@ieseg.fr", "nom": "Julie Roux"}])
    assert len(faux.appels) == 2
    assert lire_lignes(faux.appels[1]) == ["julie.roux@ieseg.fr"]  # seule la nouvelle part chez Claude
    srcs = {x["adresse"]: x["source"] for x in r2["resultats"]}
    assert srcs["marie.durand@ieseg.fr"] == "cache" and srcs["julie.roux@ieseg.fr"] == "ia"
    assert r2["cout_usd"] > 0 and r2["quota"]["utilise"] == 5


def lettres(i):
    s = ""
    i += 26
    while i:
        i, r = divmod(i, 26)
        s = chr(97 + r) + s
    return s


def lire_lignes(appel):
    return [l.split(" | ")[1] for l in appel["messages"][0]["content"].splitlines() if l.split(" | ")[0].isdigit()]


def test_reperes_transmis_a_claude(utilisateur):
    faux = brancher(FauxAnthropic())
    service.trier(utilisateur, [{"adresse": "b00654321@essec.edu"}, {"adresse": "lea.petit@essec.edu", "nom": "Léa Petit"}])
    msg = faux.appels[0]["messages"][0]["content"]
    assert "Repères" in msg and "- essec.edu : b00654321@essec.edu" in msg


def test_quota_respecte(utilisateur):
    base.modifier("Test", quota_mensuel=3)
    u = base.utilisateur_par_jeton  # noqa
    utilisateur["quota_mensuel"] = 3
    faux = brancher(FauxAnthropic())
    liste = [{"adresse": f"prenom{i}.nom@ieseg.fr"} for i in range(5)] + [{"adresse": "bde@ieseg.fr"}]
    r = service.trier(utilisateur, liste)
    src = [x["source"] for x in r["resultats"]]
    assert src.count("ia") == 3 and src.count("quota") == 2 and src.count("regle") == 1
    assert all(x["raison"].startswith("quota") for x in r["resultats"] if x["source"] == "quota")
    # quota épuisé : plus aucun appel, mais les règles et le cache marchent toujours
    r2 = service.trier(utilisateur, liste + [{"adresse": "x@gmail.com"}])
    assert len(faux.appels) == 1
    assert {x["source"] for x in r2["resultats"]} == {"cache", "quota", "regle"}
    # les « quota » ne sont pas mis en cache : ils seront triés le mois suivant
    assert "prenom4.nom@ieseg.fr" not in base.cache_lire(["prenom4.nom@ieseg.fr"])


def test_budget_mensuel_global(utilisateur, monkeypatch):
    monkeypatch.setattr(config, "BUDGET_MENSUEL_USD", 0.02)
    monkeypatch.setattr(config, "TAILLE_LOT", 2)
    monkeypatch.setattr(config, "PARALLELE", 1)
    faux = brancher(FauxAnthropic())  # chaque lot coûte 0,0168 $
    r = service.trier(utilisateur, [{"adresse": f"prenom{i}.nom@ieseg.fr"} for i in range(8)])
    src = [x["source"] for x in r["resultats"]]
    assert len(faux.appels) == 2  # 0 → 0,0168 → 0,0336 : le 3e lot est refusé
    assert src.count("ia") == 4 and src.count("budget") == 4 and r["budget_atteint"] is True
    assert base.depense_du_mois() < 0.04


def test_erreur_claude_pas_de_cache_et_message_clair(utilisateur):
    brancher(FauxAnthropic(lambda n, c, k: erreur_api(anthropic.AuthenticationError, 401)))
    r = service.trier(utilisateur, [{"adresse": "marie.durand@ieseg.fr"}, {"adresse": "bde@ieseg.fr"}])
    par = {x["adresse"]: x for x in r["resultats"]}
    assert par["marie.durand@ieseg.fr"]["groupe"] == "incertain" and par["marie.durand@ieseg.fr"]["source"] == "erreur"
    assert r["erreurs"] == ["clé API Anthropic refusée"]
    assert base.cache_lire(["marie.durand@ieseg.fr"]) == {}  # retenté au prochain tri
    assert r["quota"]["utilise"] == 0  # rien de facturé au quota
    # Clé réparée : le même contact est trié normalement
    brancher(FauxAnthropic())
    r2 = service.trier(utilisateur, [{"adresse": "marie.durand@ieseg.fr"}])
    assert r2["resultats"][0]["source"] == "ia" and r2["resultats"][0]["groupe"] == "etudiant"


def test_lot_partiel_les_oublies_restent_incertains_non_caches(utilisateur):
    def comp(n, contacts, kw):
        return reponse([{"i": 0, "groupe": "etudiant", "type": "etudiant", "raison": "ok"}])
    brancher(FauxAnthropic(comp))
    r = service.trier(utilisateur, [{"adresse": "a.b@ieseg.fr"}, {"adresse": "c.d@ieseg.fr"}])
    assert [x["groupe"] for x in r["resultats"]] == ["etudiant", "incertain"]
    assert list(base.cache_lire(["a.b@ieseg.fr", "c.d@ieseg.fr"])) == ["a.b@ieseg.fr"]


def test_gros_volume_en_lots_paralleles(utilisateur):
    faux = brancher(FauxAnthropic(latence=0.05))
    liste = [{"adresse": f"prenom{lettres(i)}.nom{lettres(i)}@ieseg.fr"} for i in range(1000)]
    r = service.trier(utilisateur, liste)
    assert len(faux.appels) == 7  # 1000 / 150 → 7 lots
    assert [x["adresse"] for x in r["resultats"]] == [c["adresse"] for c in liste]  # ordre intact malgré le parallèle
    assert r["resume"]["etudiant"] == 1000
    vus = sorted(a for appel in faux.appels for a in lire_lignes(appel))
    assert vus == sorted(c["adresse"] for c in liste)  # chaque contact envoyé une et une seule fois


def test_entrees_bizarres(utilisateur):
    brancher(FauxAnthropic())
    r = service.trier(utilisateur, [{"adresse": "  "}, {"adresse": "pas une adresse"}, {"adresse": "a@b.fr", "nom": "x" * 5000}])
    assert [x["adresse"] for x in r["resultats"]] == ["pas une adresse", "a@b.fr"]
    assert r["resultats"][0]["raison"] == "adresse invalide"
