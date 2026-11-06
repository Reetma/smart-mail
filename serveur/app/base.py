"""Base SQLite : utilisateurs (code d'accès haché + quota), consommation par mois,
et cache des tris (une adresse déjà triée ne coûte plus rien)."""

import hashlib
import json
import secrets
import sqlite3
import threading
from datetime import datetime, timedelta, timezone

from . import config

_verrou = threading.RLock()
_cnx = None

SCHEMA = """
CREATE TABLE IF NOT EXISTS utilisateurs (
  id INTEGER PRIMARY KEY,
  nom TEXT NOT NULL UNIQUE,
  jeton_hash TEXT NOT NULL UNIQUE,
  quota_mensuel INTEGER NOT NULL,
  actif INTEGER NOT NULL DEFAULT 1,
  cree TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS usage (
  utilisateur_id INTEGER NOT NULL,
  mois TEXT NOT NULL,
  contacts_total INTEGER NOT NULL DEFAULT 0,
  contacts_ia INTEGER NOT NULL DEFAULT 0,
  jetons_entree INTEGER NOT NULL DEFAULT 0,
  jetons_sortie INTEGER NOT NULL DEFAULT 0,
  cout_usd REAL NOT NULL DEFAULT 0,
  PRIMARY KEY (utilisateur_id, mois)
);
CREATE TABLE IF NOT EXISTS sessions (
  jeton_hash TEXT PRIMARY KEY,
  utilisateur_id INTEGER NOT NULL,
  cree TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS destinataires (
  utilisateur_id INTEGER NOT NULL,
  adresse TEXT NOT NULL,
  nom TEXT NOT NULL DEFAULT '',
  source TEXT NOT NULL DEFAULT '',
  ordre INTEGER NOT NULL,
  PRIMARY KEY (utilisateur_id, adresse)
);
CREATE TABLE IF NOT EXISTS recherches (
  utilisateur_id INTEGER PRIMARY KEY,
  etat TEXT NOT NULL,
  maj TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS cache (
  adresse TEXT PRIMARY KEY,
  groupe TEXT NOT NULL,
  type TEXT NOT NULL,
  raison TEXT NOT NULL,
  source TEXT NOT NULL,
  date TEXT NOT NULL
);
"""


def maintenant():
    return datetime.now(timezone.utc)


def mois_courant():
    return maintenant().strftime("%Y-%m")


def cnx():
    global _cnx
    with _verrou:
        if _cnx is None:
            config.BASE.parent.mkdir(parents=True, exist_ok=True)
            _cnx = sqlite3.connect(str(config.BASE), check_same_thread=False, isolation_level=None)
            _cnx.row_factory = sqlite3.Row
            _cnx.execute("PRAGMA journal_mode=WAL")
            _cnx.executescript(SCHEMA)
            # Colonnes ajoutées après la première version (bases existantes).
            cols = {r[1] for r in _cnx.execute("PRAGMA table_info(utilisateurs)")}
            if "email" not in cols:
                _cnx.execute("ALTER TABLE utilisateurs ADD COLUMN email TEXT")
                _cnx.execute("CREATE UNIQUE INDEX IF NOT EXISTS utilisateurs_email ON utilisateurs(email)")
            if "parametres" not in cols:
                _cnx.execute("ALTER TABLE utilisateurs ADD COLUMN parametres TEXT")
        return _cnx


def fermer():
    global _cnx
    with _verrou:
        if _cnx is not None:
            _cnx.close()
            _cnx = None


def hacher(jeton):
    return hashlib.sha256(jeton.encode()).hexdigest()


# ------------------------------------------------------------ utilisateurs
def creer_utilisateur(nom, quota):
    jeton = "sm_" + secrets.token_urlsafe(24)
    with _verrou:
        cnx().execute("INSERT INTO utilisateurs (nom, jeton_hash, quota_mensuel, cree) VALUES (?, ?, ?, ?)",
                      (nom, hacher(jeton), int(quota), maintenant().isoformat()))
    return jeton


def nouveau_code(nom):
    jeton = "sm_" + secrets.token_urlsafe(24)
    with _verrou:
        n = cnx().execute("UPDATE utilisateurs SET jeton_hash = ? WHERE nom = ?", (hacher(jeton), nom)).rowcount
    return jeton if n else None


def utilisateur_par_jeton(jeton):
    """Code d'accès créé par l'administrateur, ou jeton de session (un par appareil connecté)."""
    if not jeton:
        return None
    h = hacher(jeton)
    with _verrou:
        r = cnx().execute("SELECT * FROM utilisateurs WHERE jeton_hash = ? AND actif = 1", (h,)).fetchone()
        if not r:
            r = cnx().execute("SELECT u.* FROM sessions s JOIN utilisateurs u ON u.id = s.utilisateur_id "
                              "WHERE s.jeton_hash = ? AND u.actif = 1", (h,)).fetchone()
    return dict(r) if r else None


def connexion_par_email(email, nom="", quota=20000):
    """Trouve le compte de cette adresse ou le crée, puis ouvre une session pour CET appareil
    (les autres appareils restent connectés). Rend (jeton, utilisateur)."""
    email = email.strip().lower()
    with _verrou:
        r = cnx().execute("SELECT * FROM utilisateurs WHERE email = ?", (email,)).fetchone()
        if not r:
            cnx().execute("INSERT INTO utilisateurs (nom, email, jeton_hash, quota_mensuel, cree) VALUES (?, ?, ?, ?, ?)",
                          (nom or email, email, hacher("inutilise_" + secrets.token_urlsafe(24)), int(quota), maintenant().isoformat()))
            r = cnx().execute("SELECT * FROM utilisateurs WHERE email = ?", (email,)).fetchone()
        if not r["actif"]:
            return None, None
        jeton = "sm_" + secrets.token_urlsafe(24)
        cnx().execute("INSERT INTO sessions (jeton_hash, utilisateur_id, cree) VALUES (?, ?, ?)", (hacher(jeton), r["id"], maintenant().isoformat()))
    return jeton, dict(r)


def deconnecter(jeton):
    with _verrou:
        cnx().execute("DELETE FROM sessions WHERE jeton_hash = ?", (hacher(jeton),))


# ------------------------------------------------------------ mémoire du compte
def destinataires_ajouter(uid, personnes):
    """personnes : [{adresse, nom, source}] ; les adresses déjà présentes sont ignorées. Rend le total."""
    with _verrou:
        n = cnx().execute("SELECT COALESCE(MAX(ordre), 0) FROM destinataires WHERE utilisateur_id = ?", (uid,)).fetchone()[0]
        lignes = []
        for p in personnes:
            n += 1
            lignes.append((uid, p["adresse"].strip().lower(), (p.get("nom") or "")[:200], (p.get("source") or "")[:20], n))
        cnx().executemany("INSERT OR IGNORE INTO destinataires (utilisateur_id, adresse, nom, source, ordre) VALUES (?, ?, ?, ?, ?)", lignes)
        return cnx().execute("SELECT COUNT(*) FROM destinataires WHERE utilisateur_id = ?", (uid,)).fetchone()[0]


def destinataires_lister(uid, depart=0, taille=500):
    with _verrou:
        return [dict(r) for r in cnx().execute(
            "SELECT adresse, nom, source FROM destinataires WHERE utilisateur_id = ? ORDER BY ordre LIMIT ? OFFSET ?", (uid, taille, depart))]


def destinataires_compter(uid):
    with _verrou:
        return cnx().execute("SELECT COUNT(*) FROM destinataires WHERE utilisateur_id = ?", (uid,)).fetchone()[0]


def destinataires_vider(uid):
    with _verrou:
        cnx().execute("DELETE FROM destinataires WHERE utilisateur_id = ?", (uid,))
        cnx().execute("DELETE FROM recherches WHERE utilisateur_id = ?", (uid,))


def recherche_lire(uid):
    with _verrou:
        r = cnx().execute("SELECT etat FROM recherches WHERE utilisateur_id = ?", (uid,)).fetchone()
    return json.loads(r[0]) if r else None


def recherche_ecrire(uid, etat):
    with _verrou:
        cnx().execute("INSERT INTO recherches (utilisateur_id, etat, maj) VALUES (?, ?, ?) "
                      "ON CONFLICT(utilisateur_id) DO UPDATE SET etat = excluded.etat, maj = excluded.maj",
                      (uid, json.dumps(etat, ensure_ascii=False), maintenant().isoformat()))


def parametres_lire(uid):
    with _verrou:
        r = cnx().execute("SELECT parametres FROM utilisateurs WHERE id = ?", (uid,)).fetchone()
    return json.loads(r[0]) if r and r[0] else None


def parametres_ecrire(uid, p):
    with _verrou:
        cnx().execute("UPDATE utilisateurs SET parametres = ? WHERE id = ?", (json.dumps(p, ensure_ascii=False), uid))


def modifier(nom, **champs):
    cles = [k for k in champs if k in ("quota_mensuel", "actif")]
    if not cles:
        return 0
    with _verrou:
        return cnx().execute(
            "UPDATE utilisateurs SET " + ", ".join(k + " = ?" for k in cles) + " WHERE nom = ?",
            [champs[k] for k in cles] + [nom]).rowcount


def utilisateurs():
    with _verrou:
        return [dict(r) for r in cnx().execute("SELECT id, nom, quota_mensuel, actif, cree FROM utilisateurs ORDER BY nom")]


# ------------------------------------------------------------ consommation
def usage(uid, mois=None):
    mois = mois or mois_courant()
    with _verrou:
        r = cnx().execute("SELECT * FROM usage WHERE utilisateur_id = ? AND mois = ?", (uid, mois)).fetchone()
    return dict(r) if r else {"utilisateur_id": uid, "mois": mois, "contacts_total": 0, "contacts_ia": 0,
                              "jetons_entree": 0, "jetons_sortie": 0, "cout_usd": 0.0}


def ajouter_usage(uid, contacts_total=0, contacts_ia=0, jetons_entree=0, jetons_sortie=0, cout_usd=0.0):
    mois = mois_courant()
    with _verrou:
        cnx().execute(
            """INSERT INTO usage (utilisateur_id, mois, contacts_total, contacts_ia, jetons_entree, jetons_sortie, cout_usd)
               VALUES (?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(utilisateur_id, mois) DO UPDATE SET
                 contacts_total = contacts_total + excluded.contacts_total,
                 contacts_ia = contacts_ia + excluded.contacts_ia,
                 jetons_entree = jetons_entree + excluded.jetons_entree,
                 jetons_sortie = jetons_sortie + excluded.jetons_sortie,
                 cout_usd = cout_usd + excluded.cout_usd""",
            (uid, mois, contacts_total, contacts_ia, jetons_entree, jetons_sortie, cout_usd))


def depense_du_mois():
    with _verrou:
        r = cnx().execute("SELECT COALESCE(SUM(cout_usd), 0) FROM usage WHERE mois = ?", (mois_courant(),)).fetchone()
    return float(r[0])


def rapport(mois=None):
    mois = mois or mois_courant()
    with _verrou:
        return [dict(r) for r in cnx().execute(
            """SELECT u.nom, u.quota_mensuel, COALESCE(g.contacts_total, 0) AS contacts_total,
                      COALESCE(g.contacts_ia, 0) AS contacts_ia, COALESCE(g.cout_usd, 0) AS cout_usd
               FROM utilisateurs u LEFT JOIN usage g ON g.utilisateur_id = u.id AND g.mois = ?
               ORDER BY u.nom""", (mois,))]


# ------------------------------------------------------------ cache des tris
def cache_lire(adresses):
    if not adresses:
        return {}
    limite = (maintenant() - timedelta(days=config.CACHE_JOURS)).isoformat()
    sortie = {}
    with _verrou:
        liste = list(adresses)
        for k in range(0, len(liste), 500):
            morceau = liste[k:k + 500]
            q = "SELECT * FROM cache WHERE date >= ? AND adresse IN (" + ",".join("?" * len(morceau)) + ")"
            for r in cnx().execute(q, [limite] + morceau):
                sortie[r["adresse"]] = (r["groupe"], r["type"], r["raison"], r["source"])
    return sortie


def cache_ecrire(lignes):
    """lignes : [(adresse, groupe, type, raison, source)]"""
    if not lignes:
        return
    date = maintenant().isoformat()
    with _verrou:
        cnx().executemany(
            "INSERT INTO cache (adresse, groupe, type, raison, source, date) VALUES (?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(adresse) DO UPDATE SET groupe = excluded.groupe, type = excluded.type, "
            "raison = excluded.raison, source = excluded.source, date = excluded.date",
            [(a, g, t, r, s, date) for a, g, t, r, s in lignes])
