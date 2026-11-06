"""Tri par règles : gratuit, instantané, sûr. Ne tranche que quand l'adresse
parle d'elle-même ; tout le reste part chez Claude (ia.py).

Chaque règle rend (groupe, type, raison) ou None.
  groupe : "etudiant" | "hors" | "incertain"
  type   : "etudiant" | "perso" | "personnel" | "association" | "service" | "externe" | "inconnu"
"""

import re

# Messageries grand public : dans un annuaire d'école, ce sont des contacts perso.
MESSAGERIES_PERSO = {
    "gmail.com", "googlemail.com", "hotmail.com", "hotmail.fr", "outlook.com", "outlook.fr",
    "live.com", "live.fr", "msn.com", "yahoo.com", "yahoo.fr", "ymail.com", "icloud.com", "me.com",
    "mac.com", "orange.fr", "wanadoo.fr", "free.fr", "sfr.fr", "neuf.fr", "laposte.net", "bbox.fr",
    "aol.com", "gmx.fr", "gmx.com", "gmx.de", "proton.me", "protonmail.com", "web.de", "skynet.be",
    "telenet.be", "hotmail.be", "yahoo.be",
}

# Partie locale d'une boîte générique ou d'un service (contact@, admissions@…).
SERVICES = re.compile(
    r"^(contact|info|infos|information|admission|admissions|accueil|admin|administration|secretariat|"
    r"scolarite|communication|comm|marketing|noreply|no-reply|ne-pas-repondre|support|helpdesk|it|dsi|"
    r"rh|hr|recrutement|direction|relations?|stages?|internships?|international|exchange|career|careers|"
    r"carrieres?|entreprises?|partenariats?|presse|press|webmaster|postmaster|bibliotheque|library|"
    r"examens?|exams?|planning|registrar|finance|comptabilite|facturation|invoice|sales|hello|bonjour|"
    r"team|equipe|office|reception|events?|evenements?)\d*([._-].*)?$"
)

# Associations et clubs étudiants : ce ne sont pas des personnes.
ASSOCIATIONS = re.compile(
    r"(^|[._-])(bde|bda|bds|bdd|bdh|bdi|asso|assos|association|club|clubs|junior|je|alumni|"
    r"bureau|president|presidente|tresorier|tresoriere|vice[._-]?president|gala|wei|sport|sports)([._-]|\d|$)"
)

# Salles et ressources réservables (comptes de salle Microsoft 365).
SALLES = re.compile(r"^(salle|salles|room|rooms|amphi|amphitheatre|meeting|reunion|visio|labo)([._-]|\d|$)")

# Listes de diffusion et groupes.
LISTES = re.compile(r"^(tous|toutes|all|liste|list|listes|promo|promos|groupe|groupes|group|groups|dl|diffusion|everyone)([._-]|\d|$)")

# Identifiant étudiant : b00654321, e12345678, 20231234…
ID_ETUDIANT = re.compile(r"^[a-z]{0,2}\d{5,}$")

# Sous-domaine qui dit « étudiant » ou « personnel ».
LABELS_ETUDIANT = {"etu", "etud", "etudiant", "etudiants", "student", "students", "eleve", "eleves", "learner", "stud"}
LABELS_PERSONNEL = {"staff", "personnel", "prof", "profs", "faculty", "enseignant", "enseignants", "teacher", "teachers", "employee"}

ADRESSE = re.compile(r"^[^\s@]+@[^\s@]+\.[a-z]{2,}$")


def decouper(adresse: str):
    adresse = adresse.strip().lower()
    if not ADRESSE.match(adresse):
        return None, None
    locale, domaine = adresse.rsplit("@", 1)
    return locale, domaine


def appliquer(adresse: str, nom: str = ""):
    locale, domaine = decouper(adresse)
    if locale is None:
        return ("hors", "inconnu", "adresse invalide")
    if "#ext#" in locale or domaine.endswith(".onmicrosoft.com") and "_" in locale:
        return ("hors", "externe", "compte invité (extérieur à l'établissement)")
    if domaine in MESSAGERIES_PERSO:
        return ("hors", "perso", "messagerie personnelle (" + domaine + ")")
    if SERVICES.match(locale):
        return ("hors", "service", "boîte générique ou service")
    if SALLES.match(locale):
        return ("hors", "service", "salle ou ressource")
    if LISTES.match(locale):
        return ("hors", "service", "liste de diffusion ou groupe")
    if ASSOCIATIONS.search(locale):
        return ("hors", "association", "association ou club")
    labels = set(domaine.split(".")[:-1])
    if labels & LABELS_PERSONNEL:
        return ("hors", "personnel", "domaine du personnel")
    if ID_ETUDIANT.match(locale):
        return ("etudiant", "etudiant", "identifiant étudiant")
    if labels & LABELS_ETUDIANT:
        return ("etudiant", "etudiant", "domaine étudiant")
    return None
