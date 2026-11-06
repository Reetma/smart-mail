"""Tri par Claude des contacts que les règles n'ont pas su classer.

Un appel par lot (~150 contacts) : Claude voit tout le lot d'un coup, ce qui
lui permet de comprendre les logiques d'une école (format des adresses des
étudiants et du personnel, sigles d'associations…). La réponse est imposée au
format JSON (output_config.format), validée, et chaque contact absent ou mal
rendu retombe en « incertain » : jamais de décision inventée.
"""

import json
import time
from dataclasses import dataclass, field

import anthropic

from . import config

GROUPES = ["etudiant", "hors", "incertain"]
TYPES = ["etudiant", "perso", "personnel", "association", "service", "externe", "inconnu"]

SYSTEME = """Tu tries les contacts d'un annuaire (souvent celui d'une école ou d'une université) pour ne garder que les ÉTUDIANTS.

Pour chaque contact, réponds :
- groupe "etudiant" : c'est un étudiant ou une étudiante ;
- groupe "hors" : ce n'est pas un étudiant. Précise le type :
  "personnel" (professeur, enseignant, chercheur, membre du personnel ou de la direction),
  "association" (BDE, club, association, liste, bureau),
  "service" (boîte partagée ou générique, service, département),
  "perso" (contact privé sans lien avec l'établissement),
  "externe" (entreprise, partenaire, autre organisation) ;
- groupe "incertain" (type "inconnu") : rien ne permet de trancher.

Comment raisonner :
- Regarde le lot EN ENTIER : dans une même école, les étudiants et le personnel suivent souvent des formats d'adresse différents
  (par exemple prenom.nom@ pour les uns, p.nom@ ou nom@ pour les autres, un sous-domaine, un numéro, une année de promotion).
  Les « repères » donnés plus bas sont des adresses déjà connues comme étudiantes : sers-t'en pour reconnaître le format étudiant de chaque domaine.
- Utilise l'intitulé quand il existe (« Étudiant », « Professeur », « Programme Grande École 2027 », un département…).
- Un titre (Dr, Pr, Prof.), un intitulé de poste ou une fonction indique le personnel.
- Un nom qui n'est pas celui d'une personne (sigle, nom de club, de service, d'entreprise) n'est pas un étudiant.
- N'invente JAMAIS : un simple prénom et nom, sans autre indice, ne dit pas si la personne est étudiante ou professeure.
  Dans ce cas, réponds "incertain". Mieux vaut un incertain qu'une erreur.
- "raison" : 3 à 12 mots en français, l'indice décisif (par exemple « format prenom.nom des étudiants de ce domaine »).

Rends exactement un résultat par contact, avec son numéro i."""

SCHEMA = {
    "type": "object",
    "properties": {
        "resultats": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "i": {"type": "integer"},
                    "groupe": {"type": "string", "enum": GROUPES},
                    "type": {"type": "string", "enum": TYPES},
                    "raison": {"type": "string"},
                },
                "required": ["i", "groupe", "type", "raison"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["resultats"],
    "additionalProperties": False,
}


@dataclass
class Lot:
    resultats: dict = field(default_factory=dict)  # i -> (groupe, type, raison)
    jetons_entree: int = 0
    jetons_sortie: int = 0
    cout_usd: float = 0.0
    erreur: str = ""


def _propre(texte, n=120):
    return " ".join(str(texte or "").replace("|", "/").split())[:n]


def message_lot(contacts, reperes):
    """contacts : liste de dicts {adresse, nom, intitule} ; reperes : {domaine: [adresses étudiantes]}."""
    lignes = []
    if reperes:
        lignes.append("Repères (adresses déjà reconnues comme étudiantes, par domaine) :")
        for dom, ex in sorted(reperes.items()):
            lignes.append(f"- {dom} : " + ", ".join(ex[:5]))
        lignes.append("")
    lignes.append("Contacts à trier (i | adresse | nom | intitulé) :")
    for i, c in enumerate(contacts):
        lignes.append(f"{i} | {_propre(c['adresse'], 160)} | {_propre(c.get('nom'))} | {_propre(c.get('intitule'))}")
    return "\n".join(lignes)


def _lire(reponse, n):
    texte = next((b.text for b in reponse.content if getattr(b, "type", "") == "text"), "")
    donnees = json.loads(texte)
    sortie = {}
    for r in donnees.get("resultats", []):
        i = r.get("i")
        if isinstance(i, int) and 0 <= i < n and i not in sortie and r.get("groupe") in GROUPES:
            groupe, typ = r["groupe"], r.get("type") if r.get("type") in TYPES else "inconnu"
            # Cohérence groupe/type : pas d'« étudiant » de type « personnel ».
            if groupe == "etudiant":
                typ = "etudiant"
            elif groupe == "incertain":
                typ = "inconnu"
            elif typ in ("etudiant", "inconnu"):
                typ = "externe"
            sortie[i] = (groupe, typ, _propre(r.get("raison"), 140) or "—")
    return sortie


class ClasseurClaude:
    def __init__(self, client=None):
        # Le SDK lit ANTHROPIC_API_KEY ; il réessaie seul les 429 / 5xx / coupures.
        self.client = client or anthropic.Anthropic(max_retries=4, timeout=120.0)

    def classer(self, contacts, reperes=None) -> Lot:
        lot = Lot()
        if not contacts:
            return lot
        try:
            reponse = self.client.beta.messages.create(
                model=config.MODELE,
                max_tokens=16000,
                system=SYSTEME,
                messages=[{"role": "user", "content": message_lot(contacts, reperes or {})}],
                output_config={"effort": config.EFFORT, "format": {"type": "json_schema", "schema": SCHEMA}},
                # Si Claude refuse pour raison de sécurité (très improbable ici), Anthropic
                # relance la même requête sur le modèle de repli recommandé.
                betas=["server-side-fallback-2026-07-01"],
                fallbacks="default",
            )
        except anthropic.AuthenticationError:
            lot.erreur = "clé API Anthropic refusée"
            return lot
        except anthropic.PermissionDeniedError:
            lot.erreur = "clé API Anthropic sans autorisation"
            return lot
        except anthropic.RateLimitError:
            lot.erreur = "trop de demandes chez Anthropic, réessayez dans une minute"
            return lot
        except anthropic.APIStatusError as e:
            lot.erreur = f"erreur Anthropic {e.status_code}"
            return lot
        except anthropic.APIConnectionError:
            lot.erreur = "Anthropic injoignable"
            return lot

        u = reponse.usage
        lot.jetons_entree = int(getattr(u, "input_tokens", 0) or 0) + int(getattr(u, "cache_read_input_tokens", 0) or 0) \
            + int(getattr(u, "cache_creation_input_tokens", 0) or 0)
        lot.jetons_sortie = int(getattr(u, "output_tokens", 0) or 0)
        lot.cout_usd = config.cout(getattr(reponse, "model", config.MODELE), lot.jetons_entree, lot.jetons_sortie)

        if reponse.stop_reason == "refusal":
            lot.erreur = "Claude a refusé ce lot"
            return lot
        if reponse.stop_reason == "max_tokens" and len(contacts) > 1:
            # Réponse coupée : on recommence en deux moitiés.
            m = len(contacts) // 2
            a, b = self.classer(contacts[:m], reperes), self.classer(contacts[m:], reperes)
            lot.resultats = dict(a.resultats)
            lot.resultats.update({m + i: v for i, v in b.resultats.items()})
            lot.jetons_entree += a.jetons_entree + b.jetons_entree
            lot.jetons_sortie += a.jetons_sortie + b.jetons_sortie
            lot.cout_usd += a.cout_usd + b.cout_usd
            lot.erreur = a.erreur or b.erreur
            return lot
        try:
            lot.resultats = _lire(reponse, len(contacts))
        except (ValueError, TypeError, AttributeError):
            lot.erreur = "réponse de Claude illisible"
        return lot


class FauxClasseur:
    """Remplace Claude en démonstration et dans les tests (SMARTMAIL_FAUX_CLAUDE=1).
    Logique volontairement simple et déterministe, coût fictif."""

    MOTS_PERSONNEL = ("prof", "dr ", "dr.", "pr ", "pr.", "professeur", "enseignant", "directeur", "directrice",
                      "responsable", "chargé", "chargée", "maître", "maitre", "lecturer", "faculty", "doyen", "assistant",
                      "gestionnaire", "ingénieur", "administrator", "consultant", "invité")
    MOTS_ETUDIANT = ("étudiant", "etudiant", "student", "élève", "eleve", "promo", "programme", "bachelor", "master", "pge")

    def __init__(self):
        self.appels = 0

    def classer(self, contacts, reperes=None) -> Lot:
        self.appels += 1
        if config.FAUX_LATENCE:
            time.sleep(config.FAUX_LATENCE)  # imite la durée d'un vrai appel
        lot = Lot()
        # Prudent : ne tranche que sur un intitulé ou un titre explicite. Le format
        # des adresses, c'est le travail des règles et de la logique apprise.
        for i, c in enumerate(contacts):
            texte = (" " + (c.get("nom") or "") + " " + (c.get("intitule") or "") + " ").lower()
            if "__ia_erreur__" in c["adresse"].lower():
                continue  # simule un contact oublié par Claude
            if any(m in texte for m in self.MOTS_PERSONNEL):
                lot.resultats[i] = ("hors", "personnel", "titre ou fonction du personnel")
            elif any(m in texte for m in self.MOTS_ETUDIANT):
                lot.resultats[i] = ("etudiant", "etudiant", "intitulé étudiant")
            else:
                lot.resultats[i] = ("incertain", "inconnu", "aucun indice")
        lot.jetons_entree = 40 * len(contacts) + 900
        lot.jetons_sortie = 25 * len(contacts)
        lot.cout_usd = config.cout(config.MODELE, lot.jetons_entree, lot.jetons_sortie)
        return lot


def classeur():
    return FauxClasseur() if config.FAUX_CLAUDE else ClasseurClaude()
