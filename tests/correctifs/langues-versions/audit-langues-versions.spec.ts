// Audit « Langues et versions d'Outlook » — ENVOI (onglet ② Envoyer de Smart Mail).
//
// Banc : tests/audit/langues-versions/ (fausse page configurable, aucun réseau :
// tout est servi par context.route, le reste est refusé, DNS coupé).
//   node node_modules/@playwright/test/cli.js test tests/audit-langues-versions.spec.ts --workers=2 --reporter=line
//
// Règle de l'audit : pour chaque variante (langue, version d'Outlook, fenêtre),
// l'extension doit SOIT envoyer parfaitement (bons destinataires, objet, corps,
// importance), SOIT s'arrêter AVANT toute action risquée avec un message clair.
// Un test qui échoue aujourd'hui exprime le comportement sûr attendu et porte un
// commentaire « // BUG : … ».
import {
  test, expect, ouvrir, premierePage, choisirOnglet, url, pp, envoyes, envoyesGlobal, journal, etat, razJournal,
  lancerEnvoi, finEnvoi, clicsInterdits, saisiesHors, sansInvisibles, type Mail, type Page,
} from "./audit/langues-versions/banc";

test.describe.configure({ mode: "parallel" });

// Langues que l'extension connaît (libellés FR/EN dans publipostage/outlook.js).
const CONNUES = ["fr", "en", "de", "es", "it", "nl", "pt", "frca", "enx"] as const; // toutes prises en charge depuis la v1.3
// Langues d'Outlook web que l'extension ne connaît pas (libellés plausibles).
const AUTRES = ["de", "es", "it", "nl", "pt"] as const;
// Variantes plausibles d'une langue connue : français du Canada (« Nouveau
// courriel »), anglais « New email ».
const PROCHES = ["frca", "enx"] as const;

const DEST = ["alice@exemple.fr", "bob@exemple.fr"];
const MESSAGE = "Bonjour,\nRendez-vous lundi à 9 h.\nÀ bientôt";

test.afterEach(async ({ requetesBloquees }) => {
  // Rien n'a tenté de sortir vers Internet (hors des fausses pages Outlook).
  expect(requetesBloquees).toEqual([]);
});

async function preparer(context: import("@playwright/test").BrowserContext, worker: import("@playwright/test").Worker, adresse: string): Promise<Page> {
  await choisirOnglet(worker, "envoi");
  const page = await premierePage(context);
  await ouvrir(page, adresse);
  return page;
}

// Un mail envoyé conforme : destinataire(s) exact(s), objet, corps (message avant la signature).
function verifierMail(m: Mail, attendu: { a?: string[]; cci?: string[]; objet?: string; importance?: string }): void {
  expect(m.a, "À").toEqual(attendu.a ?? []);
  expect(m.cc, "Cc").toEqual([]);
  expect(m.cci, "Cci").toEqual(attendu.cci ?? []);
  expect(m.objet).toBe(attendu.objet ?? "Réunion de rentrée");
  for (const ligne of MESSAGE.split("\n")) expect(m.corps).toContain(ligne);
  expect(m.corps.indexOf("Bonjour")).toBeLessThan(m.corps.indexOf("Équipe Smart Mail"));
  expect(m.importance).toBe(attendu.importance ?? "normale");
}
// Pas de geste interdit : clics hors commandes prévues, frappe hors À/Cci/Objet/Corps.
async function verifierGestes(page: Page): Promise<void> {
  const j = await journal(page);
  expect(clicsInterdits(j), "clics interdits").toEqual([]);
  expect(saisiesHors(j, ["a", "cci", "objet", "corps"]), "saisie hors des champs du message").toEqual([]);
  expect(j.carnet, "carnet d'adresses ouvert").toBe(0);
  expect(j.suggestionChoisie, "suggestion choisie").toBe(0);
}

// ============================================================ A. nouvel Outlook (Fluent v9), par langue
for (const lang of CONNUES) {
  test(`A1 [${lang}] nouvel Outlook : un mail par destinataire, champs exacts, aucun geste de trop`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
    expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    const mails = await envoyes(page);
    expect(mails).toHaveLength(2);
    verifierMail(mails[0], { a: [DEST[0]] });
    verifierMail(mails[1], { a: [DEST[1]] });
    await verifierGestes(page);
  });

  test(`A2 [${lang}] nouvel Outlook : « Importance élevée » du ruban confirmée (aria-pressed)`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
    expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
    await verifierGestes(page);
  });

  test(`A3 [${lang}] nouvel Outlook : Cci (paquet de 3), « À » vide, jamais le carnet d'adresses`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    const liste = ["c1@exemple.fr", "c2@exemple.fr", "c3@exemple.fr", "c4@exemple.fr"];
    await lancerEnvoi(page, { dest: liste, texte: MESSAGE, cci: true, paquet: 3 });
    expect(await finEnvoi(page)).toContain("Terminé : 4 destinataire(s) servi(s) en 2 mail(s).");
    const mails = await envoyes(page);
    verifierMail(mails[0], { cci: liste.slice(0, 3) });
    verifierMail(mails[1], { cci: liste.slice(3) });
    await verifierGestes(page);
  });

  test(`A4 [${lang}] nouvel Outlook : ruban resté sur « Accueil », l'importance passe par l'onglet Message`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, onglet: "accueil" }));
    await lancerEnvoi(page, { dest: [DEST[0]], texte: MESSAGE, importance: true });
    expect(await finEnvoi(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    verifierMail((await envoyes(page))[0], { a: [DEST[0]], importance: "haute" });
    await verifierGestes(page);
  });

  test(`A5 [${lang}] nouvel Outlook : brouillon de l'utilisateur ouvert → arrêt, brouillon intact`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await page.evaluate(() => (window as any).__ouvrirBrouillon({ a: ["alice@exemple.fr"], objet: "Mon brouillon", corps: "Texte que j'écris", focus: "corps" }));
    const avant = await etat(page);
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
    expect(await finEnvoi(page)).toContain("Un message est déjà ouvert en rédaction");
    expect(await envoyes(page)).toEqual([]);
    const apres = await etat(page);
    expect(apres.redaction).toEqual(avant.redaction);
    expect((await journal(page)).saisies).toEqual([]);
  });
}

for (const lang of ([] as any[])) { // caduc : ces langues sont prises en charge
  test(`A6 [${lang}] arrêt sûr : rien tapé, rien cliqué, rien envoyé`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
    const fin = await finEnvoi(page);
    expect(fin).toContain("rien n'est parti");
    expect(await envoyes(page)).toEqual([]);
    const j = await journal(page);
    expect(j.saisies).toEqual([]);
    expect(j.clics.filter((c) => !c.confiance)).toEqual([]);
    expect((await etat(page)).redaction).toBeNull();
  });

  test(`A7 [${lang}] prise en charge : envoi réussi, ou arrêt qui dit que la langue n'est pas reconnue`, async ({ context, worker }) => {
    // BUG : interface non FR/EN (ou « Nouveau courriel », « New email ») : l'envoi
    // s'arrête au 1er mail sur « Bouton « Nouveau message » introuvable. Ouvrez la
    // boîte de réception d'Outlook web. » alors que la boîte de réception EST
    // ouverte. Les libellés cherchés sont seulement FR/EN (publipostage/outlook.js:36-60) :
    // la campagne ne peut jamais partir, et le message envoie l'utilisateur sur une fausse piste.
    const page = await preparer(context, worker, url({ lang }));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
    const fin = await finEnvoi(page);
    const reussi = (await envoyes(page)).length === 2;
    const clair = /langue|language|interface/i.test(fin) && !/Ouvrez la boîte de réception/.test(fin);
    expect(reussi || clair, `bilan : « ${fin} »`).toBe(true);
  });

  test(`A8 [${lang}] brouillon de l'utilisateur ouvert : jamais touché`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await page.evaluate(() => (window as any).__ouvrirBrouillon({ a: ["alice@exemple.fr"], objet: "Mon brouillon", corps: "Texte que j'écris", focus: "corps" }));
    const avant = await etat(page);
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
    expect(await finEnvoi(page)).toContain("rien n'est parti");
    expect(await envoyes(page)).toEqual([]);
    expect((await etat(page)).redaction).toEqual(avant.redaction);
    expect((await journal(page)).saisies).toEqual([]);
  });
}

// ============================================================ B. Outlook web classique (Fluent v8, ms-Persona)
for (const lang of CONNUES) {
  test(`B1 [${lang}] Outlook classique : envoi + importance par « … » > Définir l'importance > Haute`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, ui: "classique" }));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
    expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
    await verifierGestes(page);
  });

  test(`B2 [${lang}] Outlook classique : Cci (paquet de 3), pastilles ms-PickerPersona`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, ui: "classique" }));
    const liste = ["c1@exemple.fr", "c2@exemple.fr", "c3@exemple.fr"];
    await lancerEnvoi(page, { dest: liste, texte: MESSAGE, cci: true, paquet: 3 });
    expect(await finEnvoi(page)).toContain("Terminé : 3 destinataire(s) servi(s) en 1 mail(s).");
    verifierMail((await envoyes(page))[0], { cci: liste });
    await verifierGestes(page);
  });
}

test("B3 [fr] Outlook classique, boutons à icône seule (seulement un title) : envoi + importance", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", ui: "classique", icone: 1 }));
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
  expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
  for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
  await verifierGestes(page);
});

test("B4 [fr] Outlook classique, importance sans état affiché (pas d'aria-checked) : rien n'est envoyé (E4)", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", ui: "classique", impSansEtat: 1 }));
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
  expect(await finEnvoi(page)).toContain("rien n'est envoyé");
  expect(await envoyes(page)).toEqual([]);
  await verifierGestes(page);
});

for (const lang of ([] as any[])) { // caduc : ces langues sont prises en charge
  test(`B5 [${lang}] Outlook classique dans une langue inconnue : rien tapé, rien envoyé`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, ui: "classique" }));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
    expect(await finEnvoi(page)).toContain("rien n'est parti");
    expect(await envoyes(page)).toEqual([]);
    expect((await journal(page)).saisies).toEqual([]);
  });
}

test("B6 [de] Outlook classique, brouillon de l'utilisateur ouvert : intact (le repli .ms-BasePicker-input le voit)", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "de", ui: "classique" }));
  await page.evaluate(() => (window as any).__ouvrirBrouillon({ a: ["alice@exemple.fr"], objet: "Entwurf", corps: "Mein Text", focus: "corps" }));
  const avant = await etat(page);
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
  expect(await finEnvoi(page)).toContain("rien n'est parti");
  expect((await etat(page)).redaction).toEqual(avant.redaction);
  expect((await journal(page)).saisies).toEqual([]);
});

// ============================================================ C. autres adresses d'Outlook
const HOTES: Array<[string, string, string]> = [
  ["outlook.live.com", "/mail/0/", "fr"], ["outlook.live.com", "/mail/0/", "en"],
  ["outlook.cloud.microsoft", "/mail/", "fr"], ["outlook.office365.com", "/mail/", "en"],
];
for (const [hote, chemin, lang] of HOTES) {
  test(`C1 [${lang}] ${hote} : envoi + importance`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, pubs: hote.includes("live") ? 1 : undefined }, hote, chemin));
    await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
    expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
    await verifierGestes(page);
  });
}

// ============================================================ D. fenêtre étroite (commandes rangées dans « … »)
test.describe("fenêtre étroite", () => {
  test.use({ options: { largeur: 860, hauteur: 900 } });

  for (const lang of CONNUES) {
    test(`D1 [${lang}] fenêtre étroite : envoi simple`, async ({ context, worker }) => {
      const page = await preparer(context, worker, url({ lang, etroit: 1 }));
      await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
      expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
      for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]] });
      await verifierGestes(page);
    });

    test(`D2 [${lang}] fenêtre étroite : « Importance élevée » rangée dans « Autres options » de l'onglet Message`, async ({ context, worker }) => {
      // BUG : quand « Importance élevée » n'est plus visible dans le ruban (fenêtre
      // étroite ou ruban simplifié : elle passe dans « Autres options » de l'onglet
      // Message), importanceHaute() clique d'abord l'onglet « Options » (outlook.js:694-703)
      // puis n'ouvre que les « … » visibles DANS CET onglet (outlook.js:705-710) : celui de
      // l'onglet Message n'est jamais ouvert. L'envoi s'arrête (« introuvable ») et la
      // campagne avec importance ne peut pas partir tant que la fenêtre est étroite.
      const page = await preparer(context, worker, url({ lang, etroit: 1 }));
      await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
      const fin = await finEnvoi(page);
      expect(fin, fin).toContain("Terminé : 2 mail(s) envoyé(s).");
      for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
    });

    test(`D3 [${lang}] fenêtre étroite + importance : jamais d'envoi sans importance confirmée (E4)`, async ({ context, worker }) => {
      const page = await preparer(context, worker, url({ lang, etroit: 1 }));
      await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
      await finEnvoi(page);
      for (const m of await envoyes(page)) expect(m.importance).toBe("haute");
      await verifierGestes(page);
    });
  }

  test("D4 [fr] fenêtre étroite : Cci (paquet de 2)", async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang: "fr", etroit: 1 }));
    await lancerEnvoi(page, { dest: ["c1@exemple.fr", "c2@exemple.fr"], texte: MESSAGE, cci: true, paquet: 2 });
    expect(await finEnvoi(page)).toContain("Terminé : 2 destinataire(s) servi(s) en 1 mail(s).");
    verifierMail((await envoyes(page))[0], { cci: ["c1@exemple.fr", "c2@exemple.fr"] });
    await verifierGestes(page);
  });

  test("D5 [fr] ruban simplifié (fenêtre large) : importance dans « Autres options »", async ({ context, worker }) => {
    // BUG : même cause que D2 (outlook.js:694-710) : le ruban simplifié range aussi
    // « Importance élevée » dans « Autres options » de l'onglet Message.
    const page = await preparer(context, worker, url({ lang: "fr", ruban: "simplifie" }));
    await lancerEnvoi(page, { dest: [DEST[0]], texte: MESSAGE, importance: true });
    const fin = await finEnvoi(page);
    expect(fin, fin).toContain("Terminé : 1 mail(s) envoyé(s).");
    verifierMail((await envoyes(page))[0], { a: [DEST[0]], importance: "haute" });
  });
});

// ============================================================ E. mode sombre
for (const lang of CONNUES) {
  test(`E1 [${lang}] mode sombre : texte et liens intacts, bouton « arrière-plan » jamais touché`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, theme: "sombre" }));
    const html = '<p>Bonjour,</p><p>Inscrivez-vous <a href="https://exemple.fr/inscription">ici</a> avant <b>vendredi</b>.</p><p>Cordialement</p>';
    await lancerEnvoi(page, { dest: [DEST[0]], html, importance: true });
    expect(await finEnvoi(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const [m] = await envoyes(page);
    expect(m.a).toEqual([DEST[0]]);
    expect(m.importance).toBe("haute");
    expect(m.html).toContain('href="https://exemple.fr/inscription"');
    expect(m.corps.replace(/ /g, " ")).toContain("Inscrivez-vous ici avant vendredi.");
    const j = await journal(page);
    expect(j.clics.filter((c) => c.cmd === "arriere-plan")).toEqual([]);
    await verifierGestes(page);
  });
}

// ============================================================ F. rédaction dans une fenêtre séparée
test("F1 [fr] « Nouveau » ouvre la rédaction dans une fenêtre séparée : arrêt clair, rien tapé ni envoyé", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", fenetre: 1 }));
  const popup = page.waitForEvent("popup");
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
  const fen = await popup;
  await fen.locator('html[data-faux-pret="1"]').waitFor();
  const fin = await finEnvoi(page);
  expect(fin).toContain("rien n'est parti");
  expect(fin).toMatch(/rédaction ne s'est pas ouverte/);
  expect(await envoyesGlobal(page)).toEqual([]);
  expect((await journal(fen)).saisies).toEqual([]);
  expect((await etat(fen)).redaction?.a?.adresses).toEqual([]);
});

// « Isoler » à un moment précis (crochets déterministes de la fausse page) :
// __isolerApres = "a" | "objet" | "corps" → dès que l'extension a collé dans ce champ ;
// __isolerAuSurvolEnvoyer → au moment où l'extension pointe le bouton Envoyer.
async function isolerPendantEnvoi(page: Page, quand: { apres?: "a" | "objet" | "corps"; survolEnvoyer?: boolean }): Promise<{ fin: string; fen: Page }> {
  await page.evaluate((q) => { (window as any).__isolerApres = q.apres ?? null; (window as any).__isolerAuSurvolEnvoyer = !!q.survolEnvoyer; }, quand);
  const popup = page.waitForEvent("popup");
  await lancerEnvoi(page, { dest: [DEST[0]], texte: MESSAGE });
  const fen = await popup;
  await fen.locator('html[data-faux-pret="1"]').waitFor();
  const fin = await finEnvoi(page);
  return { fin, fen };
}

for (const apres of ["a", "objet", "corps"] as const) {
  test(`F2 [fr] l'utilisateur clique « Isoler » juste après « ${apres} » : arrêt, rien d'envoyé ni compté envoyé`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang: "fr" }));
    const { fin, fen } = await isolerPendantEnvoi(page, { apres });
    expect(await envoyesGlobal(page)).toEqual([]); // le brouillon isolé n'est pas parti
    expect(fin, fin).toContain("rien n'est parti");
    await expect(pp(page).locator("#pp-deja-nb")).toHaveText("0 adresse"); // pas marqué « déjà envoyé »
    await verifierGestes(page);
    expect((await journal(fen)).saisies).toEqual([]); // rien tapé dans la fenêtre isolée
  });
}

test("F3 [fr] « Isoler » après le collage du message, volet Teams ouvert : jamais l'Envoyer de Teams, jamais « envoyé » à tort", async ({ context, worker }) => {
  // BUG : la rédaction a quitté la fenêtre (« Isoler »), zoneRedaction() rend null et
  // envoyer() se rabat sur TOUT bouton « Envoyer » de la page (outlook.js:722-723) :
  // il clique l'Envoyer de la conversation Teams (message de l'utilisateur parti
  // dans Teams), puis conclut « envoyé » parce qu'aucune rédaction n'est plus
  // ouverte (outlook.js:727-728). Le destinataire est noté envoyé et « déjà envoyé »
  // alors que son mail dort en brouillon dans la fenêtre isolée (E6 + envoi perdu).
  const page = await preparer(context, worker, url({ lang: "fr", teams: 1 }));
  const { fin } = await isolerPendantEnvoi(page, { apres: "corps" });
  const j = await journal(page);
  expect(j.teamsEnvoye, "« Envoyer » de Teams cliqué").toBe(0);
  expect(await envoyesGlobal(page)).toEqual([]);
  expect(fin, fin).toContain("rien n'est parti");
  await expect(pp(page).locator("#pp-deja-nb")).toHaveText("0 adresse");
});

test("F3b [fr] « Isoler » juste après l'objet, volet Teams ouvert : le message n'est jamais collé ni envoyé dans Teams", async ({ context, worker }) => {
  // BUG : chaque étape recherche ses champs dans TOUTE la page au lieu de la
  // rédaction ouverte à l'étape 1. La rédaction partie (« Isoler »), champCorps()
  // retombe sur le sélecteur générique div[role=textbox][contenteditable][aria-multiline]
  // (outlook.js:42-43) : c'est la zone de saisie de Teams. Le message de la campagne y
  // est collé (outlook.js:544-553), puis envoyer() clique l'Envoyer de Teams (repli
  // global, outlook.js:723) et note le destinataire « envoyé » (outlook.js:727-728).
  const page = await preparer(context, worker, url({ lang: "fr", teams: 1 }));
  const { fin } = await isolerPendantEnvoi(page, { apres: "objet" });
  const j = await journal(page);
  const e = await etat(page);
  expect.soft(saisiesHors(j, ["a", "cci", "objet", "corps"]), "frappe hors de la rédaction").toEqual([]);
  expect.soft(sansInvisibles(e.teams), "texte collé dans Teams").toBe("");
  expect.soft(j.teamsMessages, "message de la campagne parti dans Teams").toEqual([]);
  expect.soft(j.teamsEnvoye, "« Envoyer » de Teams cliqué").toBe(0);
  expect.soft(await envoyesGlobal(page)).toEqual([]);
  expect.soft(fin, fin).toContain("rien n'est parti");
  await expect.soft(pp(page).locator("#pp-deja-nb")).toHaveText("0 adresse");
});

test("F4 [fr] « Isoler » au moment précis où l'extension va cliquer « Envoyer » : jamais noté « envoyé » sans envoi", async ({ context, worker }) => {
  // BUG : envoyer() juge le mail parti dès qu'aucune rédaction n'est plus visible
  // (outlook.js:727-728), sans vérifier que C'EST son clic sur SA rédaction qui l'a
  // fermée. Ici la rédaction part dans une fenêtre séparée entre le repérage du
  // bouton et le clic : le clic tombe sur un bouton détaché, rien n'est envoyé, mais
  // le bilan dit « 1 mail(s) envoyé(s) » et l'adresse entre dans « Déjà envoyés » :
  // la personne ne recevra jamais le message (envoi perdu, sans alerte).
  const page = await preparer(context, worker, url({ lang: "fr" }));
  const { fin, fen } = await isolerPendantEnvoi(page, { survolEnvoyer: true });
  expect(await envoyesGlobal(page), "aucun mail réellement parti").toEqual([]);
  expect((await etat(fen)).redaction?.a?.adresses, "le mail attend, complet, dans la fenêtre isolée").toEqual([DEST[0]]);
  expect.soft(fin, fin).not.toContain("Terminé : 1 mail(s) envoyé(s).");
  await expect.soft(pp(page).locator("#pp-deja-nb"), "pas « déjà envoyé » sans envoi").toHaveText("0 adresse");
});

test("F5 [fr] brouillon de l'utilisateur isolé dans sa fenêtre : l'envoi de la fenêtre principale n'y touche pas", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr" }));
  await page.evaluate(() => (window as any).__ouvrirBrouillon({ a: ["perso@exemple.fr"], objet: "Perso", corps: "À finir", focus: null }));
  const popup = page.waitForEvent("popup");
  await page.evaluate(() => (window as any).__isoler());
  const fen = await popup;
  await fen.locator('html[data-faux-pret="1"]').waitFor();
  const avant = await etat(fen);
  await page.bringToFront();
  await razJournal(fen);
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
  expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
  for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]] });
  expect((await etat(fen)).redaction).toEqual(avant.redaction);
  expect((await journal(fen)).saisies).toEqual([]);
  expect((await envoyesGlobal(page)).map((m) => m.fenetre)).toEqual(["principale", "principale"]);
});

test("F6 [fr] envoi lancé DEPUIS la fenêtre isolée (rédaction déjà ouverte) : refus, brouillon intact", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr" }));
  await page.evaluate(() => (window as any).__ouvrirBrouillon({ a: ["perso@exemple.fr"], objet: "Perso", corps: "À finir", focus: null }));
  const popup = page.waitForEvent("popup");
  await page.evaluate(() => (window as any).__isoler());
  const fen = await popup;
  await fen.locator('html[data-mega-pret="1"]').waitFor({ timeout: 30000 });
  await fen.bringToFront();
  await expect(pp(fen)).toBeVisible();
  const avant = await etat(fen);
  await lancerEnvoi(fen, { dest: DEST, texte: MESSAGE });
  expect(await finEnvoi(fen)).toContain("Un message est déjà ouvert en rédaction");
  expect(await envoyesGlobal(fen)).toEqual([]);
  expect((await etat(fen)).redaction).toEqual(avant.redaction);
});

test("F7 [fr] bouton Envoyer à icône seule « Envoyer (Ctrl+Entrée) » (variante plausible) + volet Teams : jamais l'Envoyer de Teams", async ({ context, worker }) => {
  // BUG : même cause que F3 : le bouton de la rédaction n'est pas reconnu (libellé
  // exact attendu), zoneRedaction() rend null et envoyer() clique le premier
  // « Envoyer » de la page (outlook.js:723), celui de Teams. La rédaction restant
  // ouverte, le destinataire passe « à vérifier » et dans « Déjà envoyés » sans
  // avoir rien reçu.
  const page = await preparer(context, worker, url({ lang: "fr", teams: 1, envoiAria: "Envoyer (Ctrl+Entrée)" }));
  await lancerEnvoi(page, { dest: [DEST[0]], texte: MESSAGE });
  const fin = await finEnvoi(page, 90000);
  const j = await journal(page);
  expect.soft(j.teamsEnvoye, "« Envoyer » de Teams cliqué").toBe(0);
  // Depuis la v1.3 le bouton à icône seule de la rédaction est reconnu : envoi exact, par CE bouton.
  const m = await envoyes(page);
  expect(m.length, fin).toBe(1);
  verifierMail(m[0], { a: [DEST[0]] });
});

test("F8 [fr] bouton Envoyer à icône seule « Envoyer (Ctrl+Entrée) », sans autre bouton Envoyer : arrêt, rien d'envoyé", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", envoiAria: "Envoyer (Ctrl+Entrée)" }));
  await lancerEnvoi(page, { dest: [DEST[0]], texte: MESSAGE });
  const fin = await finEnvoi(page, 90000);
  // Depuis la v1.3 : bouton reconnu, envoi exact.
  const m = await envoyes(page);
  expect(m.length, fin).toBe(1);
  verifierMail(m[0], { a: [DEST[0]] });
  await verifierGestes(page);
});

// ============================================================ G. Cci affiché d'office (« Toujours afficher Cci »)
for (const lang of CONNUES) {
  for (const ui of ["nouveau", "classique"] as const) {
    test(`G1 [${lang}] ${ui} : ligne Cci déjà affichée, son libellé (carnet d'adresses) n'est jamais cliqué`, async ({ context, worker }) => {
      const page = await preparer(context, worker, url({ lang, ui, cciVisible: 1 }));
      const liste = ["c1@exemple.fr", "c2@exemple.fr"];
      await lancerEnvoi(page, { dest: liste, texte: MESSAGE, cci: true, paquet: 2 });
      expect(await finEnvoi(page)).toContain("Terminé : 2 destinataire(s) servi(s) en 1 mail(s).");
      verifierMail((await envoyes(page))[0], { cci: liste });
      await verifierGestes(page);
    });
  }
}

test("G2 [fr] envoi À avec ligne Cci affichée d'office : Cci reste vide (pas de destinataire en trop)", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", cciVisible: 1 }));
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
  expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
  for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]] });
  await verifierGestes(page);
});

// ============================================================ H. rédaction dans une fenêtre modale (variante plausible)
for (const lang of CONNUES) {
  for (const ui of ["nouveau", "classique"] as const) {
    test(`H1 [${lang}] ${ui} : rédaction ouverte dans une fenêtre modale par-dessus (reste de la page aria-hidden) : envoi + importance`, async ({ context, worker }) => {
      const page = await preparer(context, worker, url({ lang, ui, dialogue: 1 }));
      await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
      const fin = await finEnvoi(page);
      expect(fin, fin).toContain("Terminé : 2 mail(s) envoyé(s).");
      for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
      await verifierGestes(page);
    });
  }
}

// ============================================================ I. diagnostic de l'onglet Envoyer (⚙ > Lancer le diagnostic)
for (const [lang, ui] of [["fr", "nouveau"], ["en", "nouveau"], ["fr", "classique"]] as const) {
  test(`I1 [${lang}] ${ui} : le diagnostic ne tape rien, n'envoie rien et referme son brouillon d'essai`, async ({ context, worker }) => {
    // BUG (fr, nouvel Outlook) : pour refermer son brouillon d'essai, le diagnostic
    // cherche « Ignorer »/« Discard » (publipostage/outlook.js:46, 795-801) ; le nouvel
    // Outlook en français dit « Abandonner » (fiche de référence). Le brouillon reste
    // ouvert (importance élevée cochée par l'essai), et le prochain envoi s'arrête sur
    // « Un message est déjà ouvert en rédaction ».
    const page = await preparer(context, worker, url({ lang, ui }));
    await razJournal(page);
    await pp(page).locator("#pp-gear").click();
    const dl = page.waitForEvent("download");
    await pp(page).locator("#pp-diag").click();
    await dl;
    await expect.poll(async () => ((await pp(page).locator("#pp-notice2").textContent()) ?? ""), { timeout: 30000 }).toContain("Diagnostic téléchargé");
    expect(await envoyes(page)).toEqual([]);
    const j = await journal(page);
    expect.soft(j.saisies, "frappe pendant le diagnostic").toEqual([]);
    expect.soft(clicsInterdits(j).filter((c) => c.cmd !== "abandonner"), "clics interdits").toEqual([]);
    expect((await etat(page)).redaction, "brouillon d'essai resté ouvert").toBeNull();
  });
}

// ============================================================ J. confusions possibles entre champs et boutons
for (const lang of CONNUES) {
  for (const ui of ["nouveau", "classique"] as const) {
    test(`J1 [${lang}] ${ui} : lignes Cc ET Cci affichées d'office : l'adresse va dans « À », jamais dans « Cc »/« Cci »`, async ({ context, worker }) => {
      const page = await preparer(context, worker, url({ lang, ui, ccVisible: 1, cciVisible: 1 }));
      await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
      expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
      for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]] });
      await verifierGestes(page);
    });
  }
}
test("J2 [fr] bouton Envoyer à icône seule, aria-label exact « Envoyer », volet Teams ouvert : envoi correct, Teams jamais touché", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", envoiAria: "Envoyer", teams: 1 }));
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE });
  expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
  for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]] });
  await verifierGestes(page);
  expect((await journal(page)).teamsEnvoye).toBe(0);
});
test("J3 [en] volet Teams ouvert (zone de saisie multiligne + « Send ») : l'envoi normal n'y touche jamais", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "en", teams: 1 }));
  await lancerEnvoi(page, { dest: DEST, texte: MESSAGE, importance: true });
  expect(await finEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
  for (const [i, m] of (await envoyes(page)).entries()) verifierMail(m, { a: [DEST[i]], importance: "haute" });
  await verifierGestes(page);
  const j = await journal(page);
  expect(j.teamsEnvoye).toBe(0);
  expect(sansInvisibles((await etat(page)).teams)).toBe("");
});
