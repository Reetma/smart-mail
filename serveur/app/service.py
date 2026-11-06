"""Le tri complet d'une liste de contacts : cache → règles → Claude, dans les
limites du quota de l'utilisateur et du budget mensuel global."""

from concurrent.futures import ThreadPoolExecutor
import threading

from . import base, config, logique as logique_mod, regles
from .ia import classeur as fabrique_classeur

# Une seule IA partagée (le client HTTP du SDK est réutilisable entre fils).
_classeur = None
_verrou_budget = threading.Lock()
# Plafond d'appels simultanés à Claude, TOUS utilisateurs confondus : évite de
# déclencher les limites de débit d'Anthropic quand 20 personnes trient en même temps.
_appels = threading.BoundedSemaphore(max(1, config.APPELS_MAX))


def classeur():
    global _classeur
    if _classeur is None:
        _classeur = fabrique_classeur()
    return _classeur


def definir_classeur(c):
    """Pour les tests : injecter un faux Claude."""
    global _classeur
    _classeur = c


def _normaliser(contacts):
    """Dédoublonne (adresse en minuscules) en gardant le premier nom/intitulé non vide."""
    vus, ordre = {}, []
    for c in contacts:
        a = str(c.get("adresse") or "").strip().lower()
        if not a:
            continue
        if a not in vus:
            vus[a] = {"adresse": a, "nom": str(c.get("nom") or "").strip()[:200], "intitule": str(c.get("intitule") or "").strip()[:200]}
            ordre.append(a)
        else:
            if not vus[a]["nom"] and c.get("nom"):
                vus[a]["nom"] = str(c["nom"]).strip()[:200]
            if not vus[a]["intitule"] and c.get("intitule"):
                vus[a]["intitule"] = str(c["intitule"]).strip()[:200]
    return [vus[a] for a in ordre]


def _reperes(etudiants_connus):
    """Adresses étudiantes connues, regroupées par domaine (5 au plus par domaine)."""
    par_dom = {}
    for a in etudiants_connus:
        dom = a.split("@", 1)[1]
        par_dom.setdefault(dom, [])
        if len(par_dom[dom]) < 5:
            par_dom[dom].append(a)
    return par_dom


def trier(utilisateur, contacts, logique=None):
    liste = _normaliser(contacts)
    res = {}  # adresse -> dict
    def poser(a, g, t, r, s):
        res[a] = {"adresse": a, "groupe": g, "type": t, "raison": r, "source": s}

    # 1. Cache
    deja = base.cache_lire([c["adresse"] for c in liste])
    for a, (g, t, r, s) in deja.items():
        poser(a, g, t, r, "cache")

    # 2. Règles
    nouvelles_regles = []
    for c in liste:
        if c["adresse"] in res:
            continue
        r = regles.appliquer(c["adresse"], c["nom"])
        if r:
            poser(c["adresse"], r[0], r[1], r[2], "regle")
            nouvelles_regles.append((c["adresse"], r[0], r[1], r[2], "regle"))
    base.cache_ecrire(nouvelles_regles)

    # 2 bis. Logique de l'école apprise sur l'échantillon : gratuite, pas mise en
    # cache (elle dépend de l'école et peut être réapprise).
    if logique is not None:
        for c in liste:
            if c["adresse"] in res:
                continue
            r = logique_mod.appliquer(c, logique)
            if r:
                poser(c["adresse"], r[0], r[1], r[2], "logique")

    # 3. Claude, pour le reste, dans la limite du quota
    reste = [c for c in liste if c["adresse"] not in res]
    u = base.usage(utilisateur["id"])
    disponible = max(0, utilisateur["quota_mensuel"] - u["contacts_ia"])
    hors_quota = reste[disponible:]
    reste = reste[:disponible]
    for c in hors_quota:
        poser(c["adresse"], "incertain", "inconnu", "quota mensuel atteint : pas trié", "quota")

    reperes = _reperes([a for a, v in res.items() if v["groupe"] == "etudiant"])
    lots = [reste[k:k + config.TAILLE_LOT] for k in range(0, len(reste), config.TAILLE_LOT)]
    erreurs, cout_total = [], 0.0
    budget_atteint = threading.Event()

    def traiter(lot):
        nonlocal cout_total
        with _verrou_budget:
            if base.depense_du_mois() >= config.BUDGET_MENSUEL_USD:
                budget_atteint.set()
                return lot, None
        with _appels:
            r = classeur().classer(lot, reperes)
        base.ajouter_usage(utilisateur["id"], contacts_ia=len(r.resultats), jetons_entree=r.jetons_entree,
                           jetons_sortie=r.jetons_sortie, cout_usd=r.cout_usd)
        with _verrou_budget:
            cout_total += r.cout_usd
        return lot, r

    if lots:
        with ThreadPoolExecutor(max_workers=max(1, config.PARALLELE)) as pool:
            for lot, r in pool.map(traiter, lots):
                if r is None:
                    for c in lot:
                        poser(c["adresse"], "incertain", "inconnu", "budget mensuel atteint : pas trié", "budget")
                    continue
                if r.erreur:
                    erreurs.append(r.erreur)
                a_cacher = []
                for i, c in enumerate(lot):
                    v = r.resultats.get(i)
                    if v:
                        poser(c["adresse"], v[0], v[1], v[2], "ia")
                        a_cacher.append((c["adresse"], v[0], v[1], v[2], "ia"))
                    else:
                        # Contact oublié ou lot en erreur : incertain, PAS mis en cache (on retentera).
                        poser(c["adresse"], "incertain", "inconnu", r.erreur or "non trié par l'IA", "erreur" if r.erreur else "ia")
                base.cache_ecrire(a_cacher)

    base.ajouter_usage(utilisateur["id"], contacts_total=len(liste))
    resultats = [res[c["adresse"]] for c in liste]
    resume = {"etudiant": 0, "hors": 0, "incertain": 0}
    for r in resultats:
        resume[r["groupe"]] += 1
    u = base.usage(utilisateur["id"])
    return {
        "resultats": resultats,
        "resume": resume,
        "quota": {"mensuel": utilisateur["quota_mensuel"], "utilise": u["contacts_ia"]},
        "cout_usd": round(cout_total, 4),
        "erreurs": sorted(set(erreurs)),
        "budget_atteint": budget_atteint.is_set(),
    }
