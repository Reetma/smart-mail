import json

import anthropic

from app import config
from app.ia import SCHEMA, SYSTEME, ClasseurClaude, message_lot
from conftest import FauxAnthropic, erreur_api, reponse

CONTACTS = [
    {"adresse": "marie.durand@ieseg.fr", "nom": "Marie Durand", "intitule": ""},
    {"adresse": "p.martin@ieseg.fr", "nom": "Pierre Martin", "intitule": ""},
    {"adresse": "lbenali@ieseg.fr", "nom": "Lina Benali", "intitule": "Professeure de finance"},
    {"adresse": "kfc@ieseg.fr", "nom": "KFC", "intitule": ""},
    {"adresse": "tom@ieseg.fr", "nom": "Tom", "intitule": ""},
]


def test_requete_envoyee_a_claude_est_conforme():
    faux = FauxAnthropic()
    lot = ClasseurClaude(client=faux).classer(CONTACTS, {"ieseg.fr": ["b.x@ieseg.fr"]})
    kw = faux.appels[0]
    assert kw["model"] == config.MODELE == "claude-opus-5-5"
    assert kw["system"] == SYSTEME
    assert kw["output_config"]["format"] == {"type": "json_schema", "schema": SCHEMA}
    assert kw["output_config"]["effort"] == "low"
    assert kw["fallbacks"] == "default" and kw["betas"] == ["server-side-fallback-2026-07-01"]
    assert "thinking" not in kw and "temperature" not in kw  # refusés par Claude Opus 5.5
    msg = kw["messages"][0]["content"]
    assert "Repères" in msg and "b.x@ieseg.fr" in msg
    assert "3 | kfc@ieseg.fr | KFC | " in msg and "4 | tom@ieseg.fr | Tom | " in msg
    assert lot.erreur == ""
    assert {i: v[0] for i, v in lot.resultats.items()} == {0: "etudiant", 1: "hors", 2: "hors", 3: "hors", 4: "incertain"}
    assert lot.resultats[2][1] == "personnel" and lot.resultats[3][1] == "association"


def test_cout_calcule_depuis_l_usage():
    lot = ClasseurClaude(client=FauxAnthropic()).classer(CONTACTS)
    assert lot.jetons_entree == 1200 and lot.jetons_sortie == 600
    assert abs(lot.cout_usd - (1200 * 4 + 600 * 20) / 1e6) < 1e-9


def test_le_message_neutralise_les_separateurs_et_les_retours_a_la_ligne():
    m = message_lot([{"adresse": "a@b.fr", "nom": "Jean | 9 | x@y.fr\nAutre", "intitule": "  Prof\t |  "}], {})
    ligne = [l for l in m.splitlines() if l.startswith("0 |")]
    assert ligne == ["0 | a@b.fr | Jean / 9 / x@y.fr Autre | Prof /"]
    assert len(m.splitlines()) == 2  # titre + une seule ligne de contact


def test_contacts_oublies_et_numeros_invalides_ne_donnent_rien():
    def comp(n, contacts, kw):
        return reponse([
            {"i": 0, "groupe": "etudiant", "type": "etudiant", "raison": "ok"},
            {"i": 0, "groupe": "hors", "type": "perso", "raison": "doublon ignoré"},
            {"i": 99, "groupe": "etudiant", "type": "etudiant", "raison": "hors plage"},
            {"i": 2, "groupe": "inconnu", "type": "etudiant", "raison": "groupe invalide"},
        ])
    lot = ClasseurClaude(client=FauxAnthropic(comp)).classer(CONTACTS)
    assert lot.resultats == {0: ("etudiant", "etudiant", "ok")}


def test_groupe_et_type_remis_en_coherence():
    def comp(n, contacts, kw):
        return reponse([
            {"i": 0, "groupe": "etudiant", "type": "personnel", "raison": "a"},
            {"i": 1, "groupe": "incertain", "type": "association", "raison": "b"},
            {"i": 2, "groupe": "hors", "type": "etudiant", "raison": "c"},
            {"i": 3, "groupe": "hors", "type": "bizarre", "raison": ""},
        ])
    r = ClasseurClaude(client=FauxAnthropic(comp)).classer(CONTACTS).resultats
    assert r[0][:2] == ("etudiant", "etudiant")
    assert r[1][:2] == ("incertain", "inconnu")
    assert r[2][:2] == ("hors", "externe")
    assert r[3] == ("hors", "externe", "—")


def test_refus_rien_n_est_classe():
    lot = ClasseurClaude(client=FauxAnthropic(lambda n, c, k: reponse([], stop_reason="refusal"))).classer(CONTACTS)
    assert lot.resultats == {} and "refusé" in lot.erreur
    assert lot.cout_usd > 0  # les jetons consommés sont comptés


def test_reponse_coupee_on_recommence_en_deux_moities():
    def comp(n, contacts, kw):
        if len(contacts) > 2:
            return reponse(texte='{"resultats": [{"i": 0, "groupe": "etu', stop_reason="max_tokens")
        return None  # réponse normale pour les petits lots
    faux = FauxAnthropic(comp)
    lot = ClasseurClaude(client=faux).classer(CONTACTS)
    assert len(lot.resultats) == 5 and lot.erreur == ""
    assert lot.resultats[0][0] == "etudiant" and lot.resultats[4][0] == "incertain"
    assert len(faux.appels) >= 3


def test_json_illisible():
    lot = ClasseurClaude(client=FauxAnthropic(lambda n, c, k: reponse(texte="Voici le tri : ..."))).classer(CONTACTS)
    assert lot.resultats == {} and "illisible" in lot.erreur


def test_erreurs_api_traduites_en_francais():
    cas = [
        (anthropic.AuthenticationError, 401, "clé API Anthropic refusée"),
        (anthropic.PermissionDeniedError, 403, "sans autorisation"),
        (anthropic.RateLimitError, 429, "trop de demandes"),
        (anthropic.InternalServerError, 500, "erreur Anthropic 500"),
        (anthropic.APIConnectionError, 0, "injoignable"),
    ]
    for classe, statut, texte in cas:
        lot = ClasseurClaude(client=FauxAnthropic(lambda n, c, k, e=erreur_api(classe, statut): e)).classer(CONTACTS)
        assert lot.resultats == {} and texte in lot.erreur, (classe, lot.erreur)


def test_lot_vide_sans_appel():
    faux = FauxAnthropic()
    assert ClasseurClaude(client=faux).classer([]).resultats == {}
    assert faux.appels == []


def test_schema_strict():
    s = SCHEMA["properties"]["resultats"]["items"]
    assert s["additionalProperties"] is False and set(s["required"]) == {"i", "groupe", "type", "raison"}
    json.dumps(SCHEMA)  # sérialisable tel quel
