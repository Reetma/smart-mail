// Audit « Relevé : variantes de la liste de suggestions du champ À (et Gmail) ».
// Partie 2 : on fait varier la liste que montre la messagerie et on vérifie les
// règles du relevé à chaque fois :
//   R1 rien de validé (ni Entrée/Tab/;/, ni clic sur une suggestion),
//   R2 rien laissé dans le champ, R3 extraction exacte (vraies adresses, bien
//   associées, rien de la boîte de réception ni d'un e-mail), R4 rien tapé
//   ailleurs, E6 aucun clic hors des commandes prévues.
// Fausses pages : audit/releve-suggestions/owa.html (nouvel Outlook : champ « À »
// contenteditable sans aria-controls, liste des messages [role=listbox] permanente,
// d'après tests/fausse-adresses/owa-reel.html et owa-fluent.html) et gmail.html.
// Quand le vrai comportement n'est pas connu (autres langues, lenteurs, groupes,
// surlignage…), la variante est PLAUSIBLE et le dit en commentaire.
import { test, expect, ao, url, ouvrir, demarrer, attendreDebut, attendreFin, etat, releve, interdites, type Banc, type Etat } from "./audit/releve-suggestions/banc";
import { PETIT, AVEC_GROUPES, adresses } from "./audit/releve-suggestions/carnet.js";

test.describe.configure({ mode: "parallel" });
test.setTimeout(420000);

const TOUS = adresses(PETIT);
const CORPS = "Bonjour,\nTexte de l'utilisateur à ne pas toucher.";
// Vrai OWA (owa-reel.html, diagnostic réel) : au départ AUCUN contact, seule la
// ligne « Afficher plus de résultats » (pied de page frère de la liste) interroge
// l'annuaire ; chaque clic déroule un lot.
const REEL = { initial: 0, lot: 3, portail: true, ariaOption: true, titleMail: true };

async function releverOWA(banc: Banc, cfg: object, o: { hote?: string; champ?: string; avant?: (p: any) => Promise<void> } = {}) {
  const { page, worker } = banc;
  await ouvrir(page, url(o.hote ?? "outlook.office.com", "/mail/inbox", cfg));
  if (o.avant) await o.avant(page);
  await page.locator(o.champ ?? "#champ-a").click({ position: { x: 8, y: 8 } });
  await demarrer(page);
  await attendreDebut(page);
  const msg = await attendreFin(page);
  return { msg, e: await etat(page), r: await releve(worker) };
}
async function releverGmail(banc: Banc, cfg: object, viewport?: { width: number; height: number }) {
  const { page, worker } = banc;
  if (viewport) await page.setViewportSize(viewport);
  await ouvrir(page, url("mail.google.com", "/mail/u/0/", cfg));
  await page.locator("#champ-to").click({ position: { x: 10, y: 10 } }); // le panneau Smart Mail recouvre le centre du champ
  await demarrer(page);
  await attendreDebut(page);
  const msg = await attendreFin(page);
  return { msg, e: await etat(page), r: await releve(worker, "gmail") };
}
// Règles de sûreté communes (R1, R2, R4, E6).
function surete(e: Etat, gmail = false) {
  expect(e.violations, "aucun geste interdit (validation, saisie ailleurs, clic hors commande)").toEqual([]);
  expect(interdites(e.keys), "jamais Entrée/Tab/;/,").toEqual([]);
  expect(e.champA, "rien laissé dans le champ").toBe("");
  expect(e.pastilles, "aucun destinataire validé").toBe(0);
  expect(e.corps.replace(/\n+/g, "\n").trim(), "corps intact").toBe(CORPS);
  if (!gmail) expect(e.objet, "objet intact").toBe("");
  expect(e.recherche, "recherche intacte").toBe("");
}

// ---------------------------------------------------------------- Outlook : bases
test.describe("Outlook — dispositions de la liste", () => {
  test("V01 vrai OWA : liste vide au départ + « Afficher plus de résultats » en pied, portail en fin de <body>", async ({ banc }) => {
    const { e, r, msg } = await releverOWA(banc, REEL);
    surete(e);
    expect(e.plusClics).toBeGreaterThan(0);
    expect(r.txt).toEqual(TOUS);
    expect(msg).toContain("Terminé");
  });
  test("V02 liste dans la rédaction (pas de portail), 5 contacts puis « Afficher plus de résultats »", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5 });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V03 10 résultats d'emblée, sans ligne « plus »", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 10, plus: { ou: "aucun", libelle: "" } });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V04 5 résultats, sans ligne « plus » (le planificateur creuse lui-même)", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, plus: { ou: "aucun", libelle: "" } });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V05 « Afficher plus de résultats » placée DANS la liste comme [role=option]", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { ...REEL, plus: { ou: "option", libelle: "Afficher plus de résultats" } });
    surete(e);
    expect(e.plusClics).toBeGreaterThan(0);
    expect(r.txt).toEqual(TOUS);
  });
  test("V06 bouton « plus » à icône seule (libellé seulement dans aria-label)", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { ...REEL, plus: { ou: "bouton-aria", libelle: "Afficher plus de résultats" } });
    surete(e);
    expect(e.plusClics).toBeGreaterThan(0);
    expect(r.txt).toEqual(TOUS);
  });
  test("V07 « Rechercher dans l'annuaire » PERMANENT (ne disparaît jamais) : pas de boucle", async ({ banc }) => {
    const { e, r, msg } = await releverOWA(banc, { cap: 5, lot: 5, plus: { ou: "pied", libelle: "Rechercher dans l'annuaire", permanent: true } });
    surete(e);
    expect(r.txt).toEqual(TOUS);
    expect(msg).toContain("Terminé");
    // garde-fou : au plus quelques clics par recherche, jamais une rafale
    expect(e.plusClics).toBeLessThanOrEqual(3 * e.typed.length);
  });
  test("V08 en-têtes de sections (« Contacts suggérés », « Annuaire ») et doublons d'une section à l'autre", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, sections: true, doublons: true });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V09 contact présent deux fois dans la même liste", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, doublons: true });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V10 badges « Externe » (adresse externe bien relevée, badge jamais pris pour une adresse)", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { ...REEL, badgeTexte: "Externe" });
    surete(e);
    expect(r.txt).toEqual(TOUS);
    expect(r.txt).toContain("baptiste.rousseau@partenaire-test.com");
  });
  test("V11 personas Fluent v8 (.ms-Persona-primaryText / secondaryText)", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { forme: "v8", cap: 5, lot: 5 });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V12 format « NOM, Prénom » dans la liste", async ({ banc }) => {
    const contacts = PETIT.map((c) => { const [p, ...n] = c.nom.split(" "); return { ...c, affichage: n.join(" ").toUpperCase() + ", " + p }; });
    const { e, r } = await releverOWA(banc, { contacts, cap: 5, lot: 5, ariaOption: true });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V13 « Démarrer » depuis la boîte de réception : ouvre le nouveau message puis relève", async ({ banc }) => {
    const { page, worker } = banc;
    await ouvrir(page, url("outlook.office.com", "/mail/inbox", { ...REEL, redaction: "absente" }));
    await expect(page.locator("#champ-a")).toHaveCount(0);
    await demarrer(page);
    await attendreDebut(page, 30000);
    await attendreFin(page);
    const e = await etat(page);
    surete(e);
    expect((await releve(worker)).txt).toEqual(TOUS);
  });
  test("V14 outlook.cloud.microsoft (même page, autre hôte)", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, REEL, { hote: "outlook.cloud.microsoft" });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
});

// ---------------------------------------------------------------- libellés « plus »
test.describe("Outlook — langue de la ligne « plus de résultats »", () => {
  for (const [id, libelle, langue] of [
    ["V15", "Rechercher dans l'annuaire", "fr"],
    ["V16", "Plus de résultats", "fr"],
    ["V17", "Show more results", "en"],
    ["V18", "Search directory", "en"],
    ["V19", "Search the directory", "en"],
  ] as const) {
    test(`${id} « ${libelle} » (${langue}) depuis une liste vide`, async ({ banc }) => {
      const { e, r } = await releverOWA(banc, { ...REEL, langue, plus: { ou: "pied", libelle } });
      surete(e);
      expect(e.plusClics).toBeGreaterThan(0);
      expect(r.txt).toEqual(TOUS);
    });
  }
  test("V20 (variante plausible : interface en allemand) « Weitere Ergebnisse anzeigen » : jamais cliquée — et le message de fin ne doit pas accuser des filtres inexistants", async ({ banc }) => {
    const { e, r, msg } = await releverOWA(banc, { ...REEL, plus: { ou: "pied", libelle: "Weitere Ergebnisse anzeigen" } });
    surete(e);
    expect(e.plusClics).toBe(0); // libellé inconnu : rien n'est cliqué (sûr)
    expect(r.txt).toEqual([]);
    // BUG : rien n'a pu être lu (Outlook ne proposait que la ligne « plus »), mais
    // le panneau conclut « Terminé, mais aucune adresse ne passe vos filtres » alors
    // qu'aucun filtre n'est réglé : l'échec est silencieux et mal expliqué (R6).
    expect(msg, "pas d'accusation de filtres quand aucun filtre n'est réglé").not.toContain("filtres");
  });
});

// ---------------------------------------------------------------- adresse peu visible
test.describe("Outlook — suggestions sans adresse visible", () => {
  test("V21 nom seul, adresse dans l'aria-label de l'option", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { forme: "nomSeul-aria", cap: 5, lot: 5 });
    surete(e);
    // BUG : lecture.js sait prendre l'adresse dans title/aria-label (« Certaines
    // versions ne mettent l'adresse que dans title/aria-label »), mais adapter-dom.js
    // ne trouve jamais la liste : boxHasMail, le repli par options et contactCount
    // ne testent que textContent. Le champ « À » réel n'a pas d'aria-controls :
    // rien n'est lu, et la fin dit « aucune adresse ne passe vos filtres ».
    expect(r.txt).toEqual(TOUS);
  });
  test("V22 nom seul, adresse dans le title de l'option", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { forme: "nomSeul-title", cap: 5, lot: 5 });
    surete(e);
    // BUG : même cause que V21 (liste jamais résolue : adresse hors textContent).
    expect(r.txt).toEqual(TOUS);
  });
  test("V23 nom seul, adresse seulement dans la carte de visite au survol (absente du DOM) : rien d'inventé", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { forme: "nomSeul-carte", cap: 5, lot: 5 });
    surete(e);
    expect(r.txt).toEqual([]);
  });
});

// ---------------------------------------------------------------- surlignage
test.describe("Outlook — partie tapée surlignée dans la liste", () => {
  test("V24 (variante plausible) surlignage <b> de la partie tapée, nom et adresse en spans Fluent : adresses EXACTES attendues", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, surligner: true });
    surete(e);
    // BUG : lecture.js (leafTexts/leafCells) lit les nœuds de texte autour du <b>
    // comme des cellules : « <b>ma</b>rtin.dupont@… » donne « rtin.dupont@… », et
    // un surlignage dans le domaine fait retomber sur textContent recollé
    // (« dupontmartin.dupont@… »). Des adresses FAUSSES entrent dans l'annuaire.
    expect(r.txt.filter((m) => !TOUS.includes(m)), "aucune adresse inventée/tronquée").toEqual([]);
    expect(r.txt).toEqual(TOUS);
  });
  test("V25 surlignage avec personas Fluent v8 (secondaryText) : adresses exactes", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { forme: "v8", cap: 5, lot: 5, surligner: true });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V26 surlignage + adresse aussi dans l'aria-label de l'option (comme le vrai Fluent) : rien de faux", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, surligner: true, ariaOption: true });
    surete(e);
    expect(r.txt.filter((m) => !TOUS.includes(m)), "aucune adresse inventée/tronquée").toEqual([]);
  });
});

// ---------------------------------------------------------------- lenteurs
test.describe("Outlook — réponses lentes ou absentes", () => {
  test("V27 aucune réponse du tout : fin propre, rien laissé, message clair", async ({ banc }) => {
    const { e, r, msg } = await releverOWA(banc, { muet: true });
    surete(e);
    expect(r.txt).toEqual([]);
    expect(msg).toContain("Terminé");
  });
  test("V28 (plausible : réseau ou annuaire lent) suggestions au bout de 2 s : le relevé doit les attendre ou dire qu'il n'a rien pu lire", async ({ banc }) => {
    const { e, r, msg } = await releverOWA(banc, { cap: 5, lot: 5, delai: 2000 });
    surete(e);
    // BUG : lireReponse abandonne chaque recherche après ~4 lectures vides
    // (≈1,5 s, rythme « normal », non réglable dans Smart Mail) ; la liste arrive
    // après. Le relevé finit « aucune adresse ne passe vos filtres » sans rien
    // signaler de la lenteur : échec silencieux.
    expect(r.txt.length > 0 || !/filtres/.test(msg), "soit des adresses, soit un message qui dit la vraie cause").toBe(true);
    expect(r.txt).toEqual(TOUS);
  });
  test("V29 suggestions au bout de 2 s, réponses périmées affichées (la liste d'avant arrive pendant la recherche suivante) : rien de faux", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, delai: 2000, perime: true });
    surete(e);
    expect(r.txt.filter((m) => !TOUS.includes(m))).toEqual([]);
  });
  test("V30 suggestions au bout de 6 s : rien laissé, rien validé, fin propre", async ({ banc }) => {
    const { e, msg } = await releverOWA(banc, { cap: 5, lot: 5, delai: 6000 });
    surete(e);
    expect(msg).toMatch(/Terminé|Arrêté/);
  });
  test("V31 (plausible) vrai OWA dont l'annuaire répond en 2 s après « Afficher plus de résultats »", async ({ banc }) => {
    const { e, r, msg } = await releverOWA(banc, { ...REEL, plus: { ou: "pied", libelle: "Afficher plus de résultats", delai: 2000 } });
    surete(e);
    // BUG : même cause que V28 — liste vide au départ, le clic « plus » part,
    // mais 3 lectures vides plus tard le moteur abandonne la recherche : rien
    // n'est relevé et la fin accuse des filtres.
    expect(r.txt.length > 0 || !/filtres/.test(msg), "soit des adresses, soit un message qui dit la vraie cause").toBe(true);
  });
});

// ---------------------------------------------------------------- listes virtualisées
test.describe("Outlook — liste virtualisée", () => {
  test("V32 liste de 10 virtualisée (4 visibles, défilement nécessaire) : rien de faux, relevé complet par les recherches suivantes", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 10, plus: { ou: "aucun", libelle: "" }, virtuel: { hauteur: 44, visibles: 4 } });
    surete(e);
    expect(r.txt.filter((m) => !TOUS.includes(m))).toEqual([]);
    expect(r.txt).toEqual(TOUS);
  });
});

// ---------------------------------------------------------------- pas des personnes
test.describe("Outlook — groupes, listes de diffusion, salles", () => {
  test("V33 (plausible) groupes / listes de diffusion / salles proposés avec leur mention visible : jamais versés dans l'annuaire d'envoi", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { carnet: "AVEC_GROUPES", cap: 5, lot: 5, ariaOption: true });
    surete(e);
    // BUG : lecture.js ne connaît que le badge « Externe » ; une suggestion
    // « Tous les étudiants M1 · Liste de distribution · tous-m1@… » est relevée
    // comme une personne et part dans l'annuaire commun, que l'onglet Envoyer
    // utilise tel quel : un seul envoi atteindrait toute la liste.
    const pasDesPersonnes = ["tous-m1@ecole-test.fr", "equipe-marketing@ecole-test.fr", "salle-a101@ecole-test.fr"];
    expect.soft(r.txt.filter((m) => pasDesPersonnes.includes(m)), "ni groupe, ni liste de diffusion, ni salle dans le relevé").toEqual([]);
    await expect.poll(async () => (await releve(banc.worker)).annuaire.length, { timeout: 10000 }).toBeGreaterThan(0);
    expect.soft((await releve(banc.worker)).annuaire.filter((m) => pasDesPersonnes.includes(m)), "ni dans l'annuaire commun").toEqual([]);
    // Et donc pas dans les destinataires de l'onglet « Envoyer » (annuaire coché par défaut) : 12 personnes.
    await ao(banc.page).locator('.mega-tab[data-onglet="envoi"]').click();
    await expect.soft(banc.page.locator("[data-pp-panneau]").locator("#pp-nb"), "destinataires de l'envoi = les 12 personnes").toHaveText("12 adresses");
  });
});

// ---------------------------------------------------------------- mauvaise liste
test.describe("Outlook — la liste lue doit être celle du champ", () => {
  test("V34 interface anglaise : boîte de réception « Message list » jamais lue", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { ...REEL, langue: "en", plus: { ou: "pied", libelle: "Show more results" }, listeMessages: { libelle: "Message list", apercus: true } });
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
  test("V35 (variante plausible : interface en allemand) boîte de réception « Nachrichtenliste » à aperçus contenant une adresse : jamais lue", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, portail: true, listeMessages: { libelle: "Nachrichtenliste", apercus: true } });
    surete(e);
    // BUG : isMsgList (adapter-dom.js) ne reconnaît la liste des messages qu'à un
    // aria-label FR/EN ; sinon resolveListbox prend la PREMIÈRE [role=listbox] à
    // adresse du document — la boîte de réception, avant le portail — et relève
    // les adresses des aperçus (noreply@…, scolarite.…) au lieu des suggestions.
    expect(r.txt.filter((m) => !TOUS.includes(m)), "rien venant de la boîte de réception").toEqual([]);
  });
  test("V36 (variante plausible) boîte de réception nommée par aria-labelledby (« Boîte de réception ») et non aria-label : jamais lue", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, portail: true, listeMessages: { libelle: "Boîte de réception", labelledby: true, apercus: true } });
    surete(e);
    // BUG : même cause que V35 (seul aria-label est regardé).
    expect(r.txt.filter((m) => !TOUS.includes(m)), "rien venant de la boîte de réception").toEqual([]);
  });
  test("V37 l'utilisateur a tapé dans Cc avant (sa liste reste montée, cachée par visibility:hidden), puis relevé dans « À »", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, portail: true, perimeCachee: true }, {
      avant: async (page) => {
        await page.locator("#champ-cc").click({ position: { x: 8, y: 8 } });
        await page.locator("#champ-cc").pressSequentially("zo", { delay: 20 });
        await page.waitForTimeout(300);
        await page.keyboard.press("Backspace"); await page.keyboard.press("Backspace");
        await page.waitForTimeout(300);
      },
    });
    surete(e);
    // BUG : visible() (adapter-dom.js) accepte un élément en visibility:hidden
    // (offsetParent non nul) : la liste périmée du Cc, plus ancienne dans le
    // portail, passe avant celle de « À » (première listbox à adresse) et est lue à
    // chaque recherche. Relevé quasi vide (« Terminé : 1 adresse »), sans alerte.
    expect(r.txt).toEqual(TOUS);
  });
  test("V38 relevé lancé depuis Cc (liste en portail) : tout est tapé dans Cc, rien dans « À »", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { ...REEL }, { champ: "#champ-cc" });
    expect(e.violations).toEqual([]);
    expect(e.typed, "rien tapé dans « À »").toEqual([]);
    expect(e.typedCc.length).toBeGreaterThan(20);
    expect(r.txt).toEqual(TOUS);
  });
  test("V39 (plausible) réponse en ligne : sous la rédaction, un e-mail de la conversation contient un lien « Voir plus » — jamais cliqué", async ({ banc }) => {
    const { e, r } = await releverOWA(banc, { cap: 5, lot: 5, conversation: { liens: ["Voir plus"] } });
    // BUG : clickMore cherche « plus de résultats » jusque dans TOUT le document
    // (NARROW_CLICK, contrainte nearField seulement) quand la liste n'a pas sa
    // propre ligne « plus » : le lien « Voir plus » de l'e-mail affiché sous la
    // réponse est cliqué, une fois par recherche (E6).
    expect(e.violations.filter((v) => v.startsWith("clic-lien-email")), "aucun lien d'e-mail cliqué").toEqual([]);
    surete(e);
    expect(r.txt).toEqual(TOUS);
  });
});

// ---------------------------------------------------------------- adresse non ASCII
test.describe("Outlook — adresses inhabituelles", () => {
  test("V40 (plausible, rare) adresse internationalisée « élodie.martin@… » : jamais relevée tronquée en « lodie.martin@… »", async ({ banc }) => {
    const contacts = [...PETIT, { nom: "Élodie Martin", mail: "élodie.martin@ecole-test.fr", poids: 70 }, { nom: "Jürgen Müller", mail: "jürgen.müller@partenaire-test.de", poids: 69, externe: true }];
    const { e, r } = await releverOWA(banc, { contacts, cap: 5, lot: 5, ariaOption: true });
    surete(e);
    // BUG : la regex d'adresse (lecture.js, moteur.js) est ASCII et non ancrée :
    // dans « élodie.martin@… » elle trouve « lodie.martin@… », dans
    // « jürgen.müller@… » elle trouve « ller@… » : des adresses d'AUTRES personnes
    // (ou inexistantes) entrent dans l'annuaire d'envoi.
    expect(r.txt.filter((m) => !TOUS.includes(m) && !m.startsWith("élodie") && !m.startsWith("jürgen")), "aucune adresse tronquée").toEqual([]);
  });
});

// ---------------------------------------------------------------- Gmail
test.describe("Gmail — champ Destinataires", () => {
  test("G01 fenêtre de rédaction, liste en portail (data-hovercard-id), région aria-live recollée", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, {});
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G02 rédaction plein écran (dialogue modal sur voile)", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { redaction: "plein-ecran" });
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G03 réponse en ligne (pas de dialogue)", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { redaction: "en-ligne", vue: "boite" });
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G04 libellé « Destinataires », champ sans aria-controls, liste dans la fenêtre", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { libelleA: "Destinataires", ariaControls: false, portail: false });
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G05 options sans data-hovercard-id (nom et adresse en blocs séparés)", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { hovercard: "aucun" });
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G06 partie tapée en gras + data-hovercard-id : adresses exactes", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { surligner: true, hovercard: "enfant" });
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G07 (variante plausible) partie tapée en gras SANS data-hovercard-id : adresses exactes attendues", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { surligner: true, hovercard: "aucun" });
    surete(e, true);
    // BUG : même cause que V24 (lecture.js leafTexts/leafCells) : « <b>ma</b>rtin.dupont@… »
    // donne « rtin.dupont@… ». Adresses fausses dans le relevé.
    expect.soft(r.txt.filter((m) => !TOUS.includes(m)), "aucune adresse inventée/tronquée").toEqual([]);
    // L'historique Gmail (base-gmail) alimente l'annuaire commun de l'envoi Outlook (mega.js, BASES).
    await expect.poll(async () => (await releve(banc.worker, "gmail")).base.length, { timeout: 10000 }).toBeGreaterThan(0);
    expect.soft((await releve(banc.worker, "gmail")).base.filter((m) => !TOUS.includes(m)), "rien de faux dans l'historique Gmail repris par l'envoi").toEqual([]);
    expect(r.txt).toEqual(TOUS);
  });
  test("G08 (plausible) fenêtre de rédaction ouverte PAR-DESSUS une lettre d'information affichée (lien « Voir plus ») : jamais cliqué", async ({ banc }) => {
    const { e, r } = await releverGmail(banc, { vue: "message", liens: ["Voir plus"] }, { width: 1440, height: 900 });
    // BUG : Gmail n'a pas de ligne « plus de résultats » : clickMore fouille tout
    // le document à chaque recherche et clique le lien « Voir plus » de l'e-mail
    // affiché sous la fenêtre de rédaction (E6), une fois par recherche.
    expect(e.violations.filter((v) => v.startsWith("clic-lien-email")), "aucun lien d'e-mail cliqué").toEqual([]);
    surete(e, true);
    expect(r.txt).toEqual(TOUS);
  });
  test("G09 (plausible) lettre affichée avec un lien « Afficher plus » / « Show more » : jamais cliqué", async ({ banc }) => {
    const { e } = await releverGmail(banc, { vue: "message", liens: ["Afficher plus", "Show more"] }, { width: 1440, height: 900 });
    // BUG : même cause que G08 (MORE_RE accepte « afficher plus » / « show more »,
    // clickMore fouille le document entier) : ≈60 clics sur le lien de l'e-mail.
    expect(e.violations.filter((v) => v.startsWith("clic-lien-email")), "aucun lien d'e-mail cliqué").toEqual([]);
  });
});
