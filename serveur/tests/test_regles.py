import pytest

from app import regles


@pytest.mark.parametrize("adresse,attendu", [
    # Messageries perso
    ("baptiste.rousseau@gmail.com", ("hors", "perso")),
    ("Marie.Durand@Hotmail.FR", ("hors", "perso")),
    ("x@icloud.com", ("hors", "perso")),
    ("jean@orange.fr", ("hors", "perso")),
    # Boîtes génériques et services
    ("contact@ieseg.fr", ("hors", "service")),
    ("admissions@essca.fr", ("hors", "service")),
    ("no-reply@ieseg.fr", ("hors", "service")),
    ("scolarite.lille@ieseg.fr", ("hors", "service")),
    ("relations-entreprises@essec.edu", ("hors", "service")),
    ("international@skema.edu", ("hors", "service")),
    ("scolarite2@ieseg.fr", ("hors", "service")),
    ("salle-a101@ieseg.fr", ("hors", "service")),
    ("room.b12@uni.ac.uk", ("hors", "service")),
    ("tous-etudiants@ieseg.fr", ("hors", "service")),
    ("promo2027@ieseg.fr", ("hors", "service")),
    ("jean.dupont_gmail.com#ext#@ieseg.onmicrosoft.com", ("hors", "externe")),
    # Associations
    ("bde@ieseg.fr", ("hors", "association")),
    ("bde.lille@ieseg.fr", ("hors", "association")),
    ("president.bds@essca.fr", ("hors", "association")),
    ("club-finance@ieseg.fr", ("hors", "association")),
    ("alumni@essec.edu", ("hors", "association")),
    # Identifiants étudiants
    ("b00654321@essec.edu", ("etudiant", "etudiant")),
    ("e12345678@univ-lille.fr", ("etudiant", "etudiant")),
    ("20231234@etu.unistra.fr", ("etudiant", "etudiant")),
    # Sous-domaines
    ("marie.durand@etu.univ-lyon1.fr", ("etudiant", "etudiant")),
    ("paul.martin@student.uliege.be", ("etudiant", "etudiant")),
    ("p.martin@staff.univ.fr", ("hors", "personnel")),
    ("p.martin@faculty.ieseg.fr", ("hors", "personnel")),
    # Invalide
    ("pas-une-adresse", ("hors", "inconnu")),
    ("a@b", ("hors", "inconnu")),
])
def test_regles_tranchent(adresse, attendu):
    r = regles.appliquer(adresse)
    assert r is not None and r[:2] == attendu, (adresse, r)
    assert r[2]  # toujours une raison lisible


@pytest.mark.parametrize("adresse", [
    # Rien dans l'adresse ne permet de trancher : c'est le travail de Claude.
    "marie.durand@ieseg.fr",
    "m.durand@ieseg.fr",
    "jdupont@essca.fr",
    "anne-laure.gautier@essca.fr",
    "contacteur.jean@ieseg.fr",   # « contact » seulement au début ET suivi d'un séparateur
    "infante.lucia@ieseg.fr",      # « info » non plus
    "allan.martin@ieseg.fr",       # « all » non plus
    "roomy.jean@ieseg.fr",         # « room » non plus
    "salleron.paul@ieseg.fr",      # ni « salle »
    "jean.club@ieseg.fr",          # « club » en fin : association… sauf que c'est un nom de famille plausible ?
])
def test_regles_laissent_passer_le_doute(adresse):
    r = regles.appliquer(adresse)
    if adresse == "jean.club@ieseg.fr":
        # Le mot « club » isolé est un signal d'association : on l'accepte (cas rarissime).
        assert r is None or r[1] == "association"
    else:
        assert r is None, (adresse, r)


def test_ne_confond_pas_les_noms_qui_contiennent_un_mot_cle():
    for a in ("jessica.bdelkader@ieseg.fr", "clubert.jean@ieseg.fr", "sportiello.marc@ieseg.fr", "asso.ufo@ieseg.fr"):
        r = regles.appliquer(a)
        if a.startswith("asso."):
            assert r and r[1] == "association"
        else:
            assert r is None, (a, r)
