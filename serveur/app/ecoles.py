"""Banc d'essai : des annuaires d'écoles fictives aux conventions DIFFÉRENTES, avec
la vérité pour chaque contact. Sert aux tests (aucun non-étudiant ne doit passer)
et au mode démonstration (/demo/annuaire?ecole=...).

Chaque contact : {adresse, nom, intitule, verite} ; verite ∈ etudiant, personnel,
service, association, perso, externe, invite, salle, liste.
"""

import random
import unicodedata

PRENOMS = ["Marie", "Lucas", "Léa", "Hugo", "Chloé", "Louis", "Emma", "Jules", "Inès", "Adam", "Lina", "Tom", "Sarah", "Nathan",
           "Zoé", "Gabriel", "Jade", "Raphaël", "Louise", "Arthur", "Alice", "Paul", "Camille", "Noah", "Manon", "Ethan", "Anna",
           "Théo", "Rose", "Mathis", "Juliette", "Sacha", "Lou", "Maël", "Eva", "Victor", "Clara", "Anne-Sophie", "Jean-Baptiste",
           "Hélène", "Zoë", "Noémie", "Loïc", "Maëlle", "Bérénice"]
NOMS = ["Durand", "Martin", "Petit", "Roux", "Morel", "Fournier", "Girard", "Bonnet", "Dupont", "Lambert", "Fontaine", "Rousseau",
        "Vincent", "Muller", "Lefèvre", "Faure", "André", "Mercier", "Blanc", "Guérin", "Boyer", "Garnier", "Chevalier", "François",
        "Legrand", "Gauthier", "Garcia", "Perrin", "Robin", "Clément", "Morin", "Nicolas", "Henry", "Roussel", "Mathieu", "Masson",
        "Dubois-Martin", "Le Gall", "N'Diaye", "Da Silva", "Nguyen", "Ben Ali", "O'Brien", "Saint-Jean"]
FIRST = ["Oliver", "Amelia", "George", "Isla", "Harry", "Ava", "Jack", "Mia", "Charlie", "Grace", "Freddie", "Lily", "Alfie", "Emily"]
LAST = ["Smith", "Jones", "Taylor", "Brown", "Williams", "Wilson", "Johnson", "Davies", "Robinson", "Wright", "Thompson", "Evans"]

POSTES = ["Professeur de finance", "Professeure de marketing", "Maître de conférences", "Assistante pédagogique",
          "Responsable des admissions", "Chargée de relations entreprises", "Directeur des programmes", "Enseignant vacataire",
          "Gestionnaire de scolarité", "Ingénieur pédagogique"]
PROGRAMMES = ["Étudiant PGE 2027", "Étudiante PGE 2026", "Bachelor 2028", "Étudiant L3 Droit", "Étudiante M1 Finance",
              "Étudiant", "Étudiante", "Programme Grande École"]


def ascii_(s):
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()
    return s.replace("'", "").replace(" ", "-")


class Fabrique:
    def __init__(self, graine):
        self.r = random.Random(graine)
        self.vus = set()
        self.sortie = []

    def unique(self, locale, dom):
        a, k = f"{locale}@{dom}", 2
        while a in self.vus:
            a = f"{locale}{k}@{dom}"
            k += 1
        self.vus.add(a)
        return a

    def ajouter(self, locale, dom, nom, intitule, verite):
        self.sortie.append({"adresse": self.unique(locale, dom), "nom": nom, "intitule": intitule, "verite": verite})

    def personne(self, anglo=False):
        if anglo:
            return self.r.choice(FIRST), self.r.choice(LAST)
        return self.r.choice(PRENOMS), self.r.choice(NOMS)

    def pieges(self, dom, n):
        """Ce qu'on trouve dans tout annuaire réel et qui n'est jamais un étudiant."""
        r = self.r
        for _ in range(n):
            x = r.random()
            p, nm = self.personne()
            if x < 0.15:
                self.ajouter(r.choice(["contact", "admissions", "scolarite", "international", "accueil", "communication"]), dom,
                             "Service", "", "service")
            elif x < 0.28:
                a = r.choice(["bde", "bds", "bda", "club-finance", "junior-conseil", "asso-humanitaire"])
                self.ajouter(a, dom, a.upper().replace("-", " "), "", "association")
            elif x < 0.40:
                self.ajouter(r.choice(["salle-a101", "salle-b204", "amphi-1", "room-c12"]), dom, "Salle " + r.choice(["A101", "B204", "Amphi 1"]),
                             "", "salle")
            elif x < 0.52:
                self.ajouter(r.choice(["tous-etudiants", "promo2027", "liste-pge", "all-students", "groupe-l3"]), dom,
                             "Liste de diffusion", "", "liste")
            elif x < 0.66:
                perso = r.choice(["gmail.com", "hotmail.fr", "outlook.fr", "icloud.com", "yahoo.fr"])
                self.ajouter(f"{ascii_(p)}.{ascii_(nm)}_{perso}#EXT#", "demoecole.onmicrosoft.com", f"{p} {nm} (Invité)", "", "invite")
            elif x < 0.80:
                self.ajouter(f"{ascii_(p)}.{ascii_(nm)}{r.randint(1, 99)}", r.choice(["gmail.com", "hotmail.fr", "orange.fr", "icloud.com"]),
                             f"{p} {nm}", "", "perso")
            else:
                self.ajouter(f"{ascii_(p)}.{ascii_(nm)}", r.choice(["deloitte-demo.com", "bnp-demo.fr", "partenaire-demo.com"]),
                             f"{p} {nm}", r.choice(["Consultant", "Chargée de recrutement", ""]), "externe")


def _melanger(f):
    f.r.shuffle(f.sortie)
    return f.sortie


# ------------------------------------------------------------ profils
def prenom_nom(n=4000, graine=7):
    """Étudiants prenom.nom@, personnel p.nom@ (intitulés parfois remplis)."""
    f, dom = Fabrique(graine), "demo-ecole.fr"
    for _ in range(int(n * 0.74)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(PROGRAMMES) if f.r.random() < 0.3 else "", "etudiant")
    for _ in range(int(n * 0.14)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)[0]}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(POSTES) if f.r.random() < 0.6 else "", "personnel")
    f.pieges(dom, int(n * 0.12))
    return _melanger(f)


def numeros(n=3000, graine=11):
    """Étudiants à numéro (b00123456@), personnel prenom.nom@ : l'inverse du profil 1."""
    f, dom = Fabrique(graine), "essec-demo.edu"
    for _ in range(int(n * 0.75)):
        p, nm = f.personne()
        f.ajouter(f"b{f.r.randint(10**7, 10**8 - 1):08d}", dom, f"{p} {nm}", "", "etudiant")
    for _ in range(int(n * 0.15)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(POSTES) if f.r.random() < 0.5 else "", "personnel")
    f.pieges(dom, int(n * 0.10))
    return _melanger(f)


def sous_domaine(n=3000, graine=13):
    """Même format prenom.nom partout ; les étudiants sont sur etu.univ-demo.fr."""
    f = Fabrique(graine)
    for _ in range(int(n * 0.78)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", "etu.univ-demo.fr", f"{p} {nm}", "", "etudiant")
    for _ in range(int(n * 0.14)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", "univ-demo.fr", f"{p} {nm}", f.r.choice(POSTES) if f.r.random() < 0.4 else "", "personnel")
    f.pieges("univ-demo.fr", int(n * 0.08))
    return _melanger(f)


def intitules_seuls(n=3000, graine=17):
    """Tout le monde en prenom.nom sur le même domaine : seul l'intitulé permet de trancher."""
    f, dom = Fabrique(graine), "ecole-unique.fr"
    for _ in range(int(n * 0.72)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(PROGRAMMES) if f.r.random() < 0.7 else "", "etudiant")
    for _ in range(int(n * 0.18)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(POSTES) if f.r.random() < 0.8 else "", "personnel")
    f.pieges(dom, int(n * 0.10))
    return _melanger(f)


def nom_prenom(n=3000, graine=19):
    """Étudiants nom.prenom@, personnel prenom.nom@ : deux formats à point, piégeux."""
    f, dom = Fabrique(graine), "ecole-np.fr"
    for _ in range(int(n * 0.75)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(nm)}.{ascii_(p)}", dom, f"{p} {nm}", f.r.choice(PROGRAMMES) if f.r.random() < 0.25 else "", "etudiant")
    for _ in range(int(n * 0.15)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(POSTES) if f.r.random() < 0.6 else "", "personnel")
    f.pieges(dom, int(n * 0.10))
    return _melanger(f)


def multi_domaines(n=3000, graine=23):
    """Deux campus et un ancien domaine ; étudiants prenom.nom, personnel p.nom partout."""
    f = Fabrique(graine)
    doms = ["campus-lille.demo.fr", "campus-paris.demo.fr"]
    for _ in range(int(n * 0.74)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", f.r.choice(doms), f"{p} {nm}", f.r.choice(PROGRAMMES) if f.r.random() < 0.3 else "", "etudiant")
    for _ in range(int(n * 0.16)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)[0]}.{ascii_(nm)}", f.r.choice(doms + ["ancien-demo.fr"]), f"{p} {nm}",
                  f.r.choice(POSTES) if f.r.random() < 0.6 else "", "personnel")
    f.pieges(doms[0], int(n * 0.10))
    return _melanger(f)


def anglophone(n=3000, graine=29):
    """École anglophone : Student / Lecturer, firstname.lastname@ pour tous, staff avec intitulé."""
    f, dom = Fabrique(graine), "uni-demo.ac.uk"
    for _ in range(int(n * 0.75)):
        p, nm = f.personne(anglo=True)
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(["Student", "Undergraduate Student", "MSc Student", ""]), "etudiant")
    for _ in range(int(n * 0.15)):
        p, nm = f.personne(anglo=True)
        f.ajouter(f"{ascii_(p)[0]}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(["Lecturer", "Senior Lecturer", "Professor", "Administrator", ""]),
                  "personnel")
    f.pieges(dom, int(n * 0.10))
    return _melanger(f)


def pieges(n=1500, graine=31):
    """Annuaire très « sale » : un tiers de pièges, des noms composés, accents, homonymes."""
    f, dom = Fabrique(graine), "demo-pieges.fr"
    for _ in range(int(n * 0.5)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(PROGRAMMES) if f.r.random() < 0.35 else "", "etudiant")
    for _ in range(int(n * 0.15)):
        p, nm = f.personne()
        f.ajouter(f"{ascii_(p)[0]}.{ascii_(nm)}", dom, f"{p} {nm}", f.r.choice(POSTES) if f.r.random() < 0.6 else "", "personnel")
    f.pieges(dom, int(n * 0.35))
    return _melanger(f)


PROFILS = {
    "prenom-nom": prenom_nom, "numeros": numeros, "sous-domaine": sous_domaine, "intitules-seuls": intitules_seuls,
    "nom-prenom": nom_prenom, "multi-domaines": multi_domaines, "anglophone": anglophone, "pieges": pieges,
}
