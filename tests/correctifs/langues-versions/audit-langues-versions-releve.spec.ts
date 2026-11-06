// Audit « Langues et versions d'Outlook » — RELEVÉ D'ADRESSES (onglet ① Relever).
//
// Banc : tests/audit/langues-versions/ (fausse page configurable, aucun réseau).
//   node node_modules/@playwright/test/cli.js test tests/audit-langues-versions-releve.spec.ts --workers=2 --reporter=line
//
// Le relevé tape des préfixes dans le champ « À » et lit les suggestions. Règles :
// R1 ne jamais valider, R2 ne rien laisser dans le champ, R3 que de vraies
// adresses de la liste de suggestions (rien de la boîte de réception), R4 ne
// jamais taper ailleurs que dans le champ destinataire, R5 ne jamais effacer ce
// que l'utilisateur avait mis. Le relevé n'a pas de mode accéléré : chaque test
// laisse tourner le vrai rythme quelques secondes, puis met en pause.
import {
  test, expect, ouvrir, premierePage, choisirOnglet, url, ao, journal, etat, razJournal,
  demarrerReleve, adressesReleve, noticeReleve, pauseReleve, pastilleReleve, attendreJusqua,
  saisiesHors, touchesValidation, clicsInterdits, sansInvisibles, ANNUAIRE, pp, envoyesGlobal, remplirEnvoi, type Page,
} from "./audit/langues-versions/banc";

test.describe.configure({ mode: "parallel" });

const CONNUES = ["fr", "en", "de", "es", "it", "nl", "pt", "frca", "enx"] as const; // toutes prises en charge depuis la v1.3
const AUTRES = ["de", "es", "it", "nl", "pt"] as const;
const TOUTES = Array.from(new Set([...CONNUES, ...AUTRES])) as any[];
const BOITE = ["noreply@email.apple.com", "learn@mails.microsoft.com", "infos@contoso-voyages.example"]; // adresses visibles dans la boîte de réception

test.afterEach(async ({ requetesBloquees }) => { expect(requetesBloquees).toEqual([]); });

async function preparer(context: import("@playwright/test").BrowserContext, worker: import("@playwright/test").Worker, adresse: string): Promise<Page> {
  await choisirOnglet(worker, "adresses");
  const page = await premierePage(context);
  await ouvrir(page, adresse);
  await expect(ao(page).locator("#ao-run")).toBeVisible();
  return page;
}
// Un message ouvert (nouveau, vide) avec le curseur dans « À ».
async function messageOuvert(page: Page, o: Record<string, unknown> = {}): Promise<void> {
  await page.evaluate((x) => (window as any).__ouvrirBrouillon(x), { focus: "a", ...o });
}
// Laisse tourner le relevé jusqu'à « ok » (ou « ms » au plus), puis le met en pause.
async function laisserTourner(page: Page, ms: number, ok: (a: string[]) => boolean = () => false): Promise<string[]> {
  await expect(pastilleReleve(page)).toBeVisible({ timeout: 20000 });
  await attendreJusqua(() => adressesReleve(page), ok, ms, 300);
  await pauseReleve(page);
  return adressesReleve(page);
}

// ============================================================ RA. nouvel Outlook, curseur dans « À », par langue
for (const lang of TOUTES) {
  test(`RA1 [${lang}] curseur dans « À » : ne tape que dans « À », ne valide rien, ne laisse rien`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await messageOuvert(page);
    await demarrerReleve(page);
    await laisserTourner(page, 12000);
    const j = await journal(page);
    expect(saisiesHors(j, ["a"]), "frappe hors de « À »").toEqual([]);
    expect(j.saisies.length, "le relevé a bien tapé dans « À »").toBeGreaterThan(0);
    expect(touchesValidation(j), "Entrée/Tab/;/, envoyés").toEqual([]);
    expect(j.suggestionChoisie, "suggestion cliquée (validation)").toBe(0);
    expect(j.carnet).toBe(0);
    expect(clicsInterdits(j), "clic hors de la liste de suggestions (boîte, carnet, Teams…)").toEqual([]);
    const e = await etat(page);
    expect(e.redaction?.a?.pastilles, "destinataire validé").toBe(0);
    expect(sansInvisibles(e.redaction?.a?.libre), "texte laissé dans « À »").toBe("");
  });

  test(`RA2 [${lang}] curseur dans « À » : rien n'est pris dans la boîte de réception (R3)`, async ({ context, worker }) => {
    // BUG (de, es, it, nl, pt) : la liste des messages est une [role=listbox] dont un
    // aperçu contient une adresse ; elle n'est écartée que par son libellé FR/EN
    // (« Liste des messages », « Message list »… : adresses/adapter-dom.js:145-148). Dans
    // une autre langue (« Nachrichtenliste », « Lista de mensajes »…), resolveListbox()
    // la prend pour la liste de suggestions (adapter-dom.js:160-161) : les adresses de la
    // boîte de réception (noreply@email.apple.com…) entrent dans l'annuaire commun, donc
    // dans les destinataires de l'onglet Envoyer.
    const page = await preparer(context, worker, url({ lang }));
    await messageOuvert(page);
    await demarrerReleve(page);
    const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3 || a.some((x) => BOITE.includes(x)));
    expect(trouvees.filter((x) => BOITE.includes(x)), "adresses de la boîte de réception relevées").toEqual([]);
    expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x)), "adresses étrangères au carnet").toEqual([]);
  });

  test(`RA3 [${lang}] curseur dans « À » : « Afficher plus de résultats » est déroulé, l'annuaire est relevé`, async ({ context, worker }) => {
    // BUG (de, es, it, nl, pt) : la ligne « plus de résultats / rechercher dans
    // l'annuaire » n'est reconnue qu'en FR/EN (MORE_RE, adresses/adapter-dom.js:178).
    // Sans ce clic, le nouvel Outlook ne montre aucun contact de l'annuaire : le relevé
    // tourne à vide (ou ne relève que ce qui traîne ailleurs) sans jamais dire pourquoi.
    const page = await preparer(context, worker, url({ lang }));
    await messageOuvert(page);
    await demarrerReleve(page);
    const trouvees = await laisserTourner(page, 15000, (a) => a.filter((x) => ANNUAIRE.adresses.includes(x)).length >= 3);
    const j = await journal(page);
    expect(j.clics.filter((c) => c.cmd === "plus-resultats" && !c.confiance).length, "clics sur « plus de résultats »").toBeGreaterThan(0);
    expect(trouvees.filter((x) => ANNUAIRE.adresses.includes(x)).length, `contacts relevés : ${trouvees.join(", ")}`).toBeGreaterThanOrEqual(3);
  });
}

// ============================================================ RB. « Démarrer » depuis la boîte de réception
for (const lang of CONNUES) {
  test(`RB1 [${lang}] « Démarrer » sans message ouvert : ouvre un nouveau message et relève dans son « À »`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await demarrerReleve(page);
    const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3);
    expect((await etat(page)).redaction, "nouveau message ouvert").not.toBeNull();
    const j = await journal(page);
    expect(saisiesHors(j, ["a"])).toEqual([]);
    expect(trouvees.length).toBeGreaterThanOrEqual(3);
    expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
  });
}
for (const lang of ([] as any[])) { // caduc : ces langues sont prises en charge
  test(`RB2 [${lang}] « Démarrer » sans message ouvert : arrêt sûr (rien tapé nulle part)`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang }));
    await demarrerReleve(page);
    await expect.poll(() => noticeReleve(page), { timeout: 20000 }).not.toBe("");
    expect((await journal(page)).saisies).toEqual([]);
    expect(await pastilleReleve(page).isVisible()).toBe(false);
  });

  test(`RB3 [${lang}] « Démarrer » sans message ouvert : le message d'arrêt dit que la langue n'est pas reconnue`, async ({ context, worker }) => {
    // BUG : « Bouton « Nouveau message » introuvable. Ouvrez la boîte de réception
    // d'Outlook web. » alors que la boîte de réception est ouverte : seul le libellé
    // FR/EN du bouton est cherché (publipostage/outlook.js:36-37, 264-265). Le relevé est
    // impossible à démarrer ainsi, et l'utilisateur est envoyé sur une fausse piste.
    const page = await preparer(context, worker, url({ lang }));
    await demarrerReleve(page);
    await expect.poll(() => noticeReleve(page), { timeout: 20000 }).not.toBe("");
    const n = await noticeReleve(page);
    expect(/langue|language|interface/i.test(n) && !/Ouvrez la boîte de réception/.test(n), `notice : « ${n} »`).toBe(true);
  });
}

// ============================================================ RC. curseur ailleurs que dans un champ destinataire
// L'utilisateur écrivait dans son message (objet, corps) ou venait de chercher un
// mail, puis clique « Démarrer » : le relevé ne doit JAMAIS taper ni effacer là.
async function curseurPuisDemarrer(page: Page, ou: "objet" | "corps" | "recherche"): Promise<{ avant: string; apres: string; saisies: string[] }> {
  if (ou === "recherche") {
    await page.locator('[data-champ="recherche"]').fill("facture juin");
    await page.evaluate(() => (window as any).__placerCurseur("recherche", "fin"));
  } else {
    await messageOuvert(page, { a: [], objet: "Compte rendu du conseil", corps: "Bonjour à toutes et à tous, voici le compte rendu promis.", focus: ou, caret: ou === "corps" ? "texte" : "fin" });
  }
  const lire = async () => { const e = await etat(page); return ou === "recherche" ? e.recherche : ou === "objet" ? e.redaction!.objet : e.redaction!.corpsTexte; };
  const avant = await lire();
  await demarrerReleve(page);
  // Le défaut, s'il existe, se voit dès la 1re recherche (effacement puis frappe) ;
  // on laisse alors tourner 5 s de plus pour mesurer le dégât, puis pause.
  const change = await attendreJusqua(lire, (v) => v !== avant, 8000, 200);
  if (change !== avant) await new Promise((r) => setTimeout(r, 5000));
  await pauseReleve(page);
  const apres = await lire();
  const j = await journal(page);
  return { avant, apres, saisies: j.saisies.map((s) => s.champ) };
}
for (const lang of TOUTES) {
  test(`RC1 [${lang}] curseur dans l'Objet de son message, puis « Démarrer » : l'objet n'est ni effacé ni modifié`, async ({ context, worker }) => {
    // BUG (de, es, it, nl, pt) : champActif() (adresses/panel.js:465-471) prend le champ
    // du curseur s'il n'est pas étiqueté search/recherch/objet/subject : « Betreff »,
    // « Asunto », « Oggetto », « Onderwerp », « Assunto » passent. Le relevé efface
    // jusqu'à 200 caractères de l'objet (1er effacement, adapter-dom.js:103) puis y tape
    // ses préfixes (R4, R5).
    const page = await preparer(context, worker, url({ lang }));
    const r = await curseurPuisDemarrer(page, "objet");
    expect.soft(r.apres, "objet après « Démarrer »").toBe(r.avant);
    expect(r.saisies.filter((c) => c === "objet").length, "frappes/effacements dans l'objet").toBe(0);
  });

  test(`RC2 [${lang}] curseur dans le corps de son message, puis « Démarrer » : le texte n'est ni effacé ni modifié`, async ({ context, worker }) => {
    // BUG (toutes langues, FR compris) : champActif() (adresses/panel.js:465-471) n'écarte
    // pas le corps du message (aucun motif corps/body/message), armHere() l'arme sans
    // condition (armement.js:132-145) : le 1er effacement retire jusqu'à 200 caractères
    // du texte de l'utilisateur avant le curseur (adapter-dom.js:99-112), puis chaque
    // recherche tape un préfixe et en efface deux de plus (core/moteur.js:510).
    const page = await preparer(context, worker, url({ lang }));
    const r = await curseurPuisDemarrer(page, "corps");
    expect.soft(r.apres, "corps après « Démarrer »").toBe(r.avant);
    expect(r.saisies.filter((c) => c === "corps").length, "frappes/effacements dans le corps").toBe(0);
  });
}
for (const lang of TOUTES) {
  test(`RC3 [${lang}] curseur dans la recherche (barre hors repère, variante plausible), puis « Démarrer » : rien n'y est tapé`, async ({ context, worker }) => {
    // BUG (de, es, it, nl, pt) : hors repère role=search/banner/header, la barre de
    // recherche n'est écartée que par les mots « search »/« recherch » de son libellé
    // (adresses/panel.js:470, adapter-dom.js:40-45) : « Suchen », « Buscar », « Cerca »,
    // « Zoeken », « Pesquisar » passent, et le relevé tape dans la recherche.
    const page = await preparer(context, worker, url({ lang, landmark: 0 }));
    const r = await curseurPuisDemarrer(page, "recherche");
    expect.soft(r.apres, "recherche après « Démarrer »").toBe(r.avant);
    expect(r.saisies.filter((c) => c === "recherche").length, "frappes/effacements dans la recherche").toBe(0);
  });
}
test("RC4 [de] curseur dans la recherche dans un repère role=search : rien n'y est tapé", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "de" }));
  const r = await curseurPuisDemarrer(page, "recherche");
  expect(r.saisies.filter((c) => c === "recherche").length).toBe(0);
  expect(r.apres).toBe(r.avant);
});

// ============================================================ RD. « À » déjà rempli par l'utilisateur
for (const lang of CONNUES) {
  test(`RD1 [${lang}] nouvel Outlook : « À » contient déjà un destinataire validé (pastille), curseur après lui : la pastille reste`, async ({ context, worker }) => {
    // BUG : dans le nouvel Outlook, les pastilles sont DANS le champ éditable ; le 1er
    // effacement du passage (adresses/adapter-dom.js:103, jusqu'à 200 retours arrière
    // tant que textContent n'est pas vide, adapter-dom.js:98 et 104-105) supprime les
    // destinataires que l'utilisateur avait déjà mis (R5). Le relevé finit en disant
    // « champ « À » vide ».
    const page = await preparer(context, worker, url({ lang }));
    await messageOuvert(page, { a: ["alice@exemple.fr"], objet: "Brouillon", corps: "Texte", focus: "a", caret: "fin" });
    await demarrerReleve(page);
    await attendreJusqua(async () => (await etat(page)).redaction?.a?.pastilles ?? 0, (n) => n === 0, 8000, 200);
    await pauseReleve(page);
    const e = await etat(page);
    expect(e.redaction?.a?.adresses, "destinataire de l'utilisateur").toEqual(["alice@exemple.fr"]);
  });
}
test("RD2 [fr] nouvel Outlook : texte que l'utilisateur tapait dans « À » (non validé), curseur après : il n'est pas effacé", async ({ context, worker }) => {
  // BUG : même cause que RD1 (adresses/adapter-dom.js:103) : le 1er effacement vide
  // le champ, ce que l'utilisateur y avait tapé est perdu (R5).
  const page = await preparer(context, worker, url({ lang: "fr" }));
  await messageOuvert(page, { focus: null });
  await page.evaluate(() => { const a = document.querySelector('[data-champ="a"]') as HTMLElement; a.textContent = "jean.dup"; });
  await page.evaluate(() => (window as any).__placerCurseur("a", "fin"));
  await demarrerReleve(page);
  await attendreJusqua(async () => sansInvisibles((await etat(page)).redaction?.a?.libre), (v) => v !== "jean.dup", 8000, 200);
  await pauseReleve(page);
  expect(sansInvisibles((await etat(page)).redaction?.a?.libre)).toBe("jean.dup");
});
test("RD3 [fr] Outlook classique : pastilles hors du champ (ms-PickerPersona) : elles restent", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", ui: "classique" }));
  await messageOuvert(page, { a: ["alice@exemple.fr"], objet: "Brouillon", corps: "Texte", focus: "a", caret: "fin" });
  await demarrerReleve(page);
  await page.waitForTimeout(3000);
  if (await pastilleReleve(page).isVisible()) await laisserTourner(page, 8000); // sinon : refus prudent (champ non vide)
  expect((await etat(page)).redaction?.a?.adresses).toEqual(["alice@exemple.fr"]);
});
test("RD4 [fr] nouvel Outlook : brouillon avec destinataire, curseur ailleurs (« Démarrer » prend le « À » du message ouvert) : la pastille reste", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr" }));
  await messageOuvert(page, { a: ["alice@exemple.fr"], objet: "Brouillon", corps: "Texte", focus: null });
  await demarrerReleve(page);
  await page.waitForTimeout(3000);
  if (await pastilleReleve(page).isVisible()) await laisserTourner(page, 10000); // sinon : refus prudent (message de l'utilisateur)
  const e = await etat(page);
  expect(e.redaction?.a?.adresses).toEqual(["alice@exemple.fr"]);
  expect(sansInvisibles(e.redaction?.a?.libre)).toBe("");
});

// ============================================================ RE. versions d'Outlook
for (const lang of CONNUES) {
  test(`RE1 [${lang}] Outlook classique (ms-BasePicker, aria-controls) : relève l'annuaire, rien d'autre`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, ui: "classique" }));
    await messageOuvert(page);
    await demarrerReleve(page);
    const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 5);
    expect(trouvees.length).toBeGreaterThanOrEqual(5);
    expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
    const j = await journal(page);
    expect(saisiesHors(j, ["a"])).toEqual([]);
    expect(touchesValidation(j)).toEqual([]);
    expect((await etat(page)).redaction?.a?.pastilles).toBe(0);
  });
}
test("RE2 [fr] Outlook classique : « Démarrer » depuis la boîte de réception (« Nouveau message »)", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", ui: "classique" }));
  await demarrerReleve(page);
  const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3);
  expect(trouvees.length).toBeGreaterThanOrEqual(3);
  expect(saisiesHors(await journal(page), ["a"])).toEqual([]);
});
test("RE3 [de] Outlook classique, message ouvert, curseur ailleurs : relève dans « An » (repli .ms-BasePicker-input), mais sans « Verzeichnis durchsuchen »", async ({ context, worker }) => {
  // BUG : même cause que RA3 (MORE_RE FR/EN, adresses/adapter-dom.js:178) : en classique
  // aussi, « Verzeichnis durchsuchen » n'est jamais cliqué ; seules les 2 premières
  // suggestions du cache sont lues pour chaque recherche.
  const page = await preparer(context, worker, url({ lang: "de", ui: "classique" }));
  await messageOuvert(page, { focus: null });
  await demarrerReleve(page);
  await laisserTourner(page, 12000, (a) => a.length >= 5);
  const j = await journal(page);
  expect.soft(saisiesHors(j, ["a"]), "frappe hors de « An »").toEqual([]);
  expect(j.clics.filter((c) => c.cmd === "plus-resultats" && !c.confiance).length, "clics « Verzeichnis durchsuchen »").toBeGreaterThan(0);
});
for (const [hote, chemin] of [["outlook.live.com", "/mail/0/"], ["outlook.cloud.microsoft", "/mail/"]] as const) {
  test(`RE4 [fr] ${hote} : relève l'annuaire, rien de la boîte (annonce comprise)`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang: "fr", pubs: 1 }, hote, chemin));
    await messageOuvert(page);
    await demarrerReleve(page);
    const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3);
    expect(trouvees.length).toBeGreaterThanOrEqual(3);
    expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
  });
}
for (const lang of CONNUES) {
  test(`RE5 [${lang}] mode sombre : relève l'annuaire, ne tape que dans « À »`, async ({ context, worker }) => {
    const page = await preparer(context, worker, url({ lang, theme: "sombre" }));
    await messageOuvert(page);
    await demarrerReleve(page);
    const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3);
    expect(trouvees.length).toBeGreaterThanOrEqual(3);
    expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
    expect(saisiesHors(await journal(page), ["a"])).toEqual([]);
  });
}
test.describe("fenêtre étroite", () => {
  test.use({ options: { largeur: 860, hauteur: 900 } });
  for (const lang of CONNUES) {
    test(`RE6 [${lang}] fenêtre étroite : relève l'annuaire, ne tape que dans « À »`, async ({ context, worker }) => {
      const page = await preparer(context, worker, url({ lang, etroit: 1 }));
      await messageOuvert(page);
      await demarrerReleve(page);
      const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3);
      expect(trouvees.length).toBeGreaterThanOrEqual(3);
      expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
      expect(saisiesHors(await journal(page), ["a"])).toEqual([]);
    });
  }
});

// ============================================================ RF. fenêtres séparées
test("RF1 [fr] rédaction isolée (fenêtre /mail/deeplink/compose) : « Démarrer » dans cette fenêtre relève dans son « À »", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr" }));
  await messageOuvert(page, { focus: null });
  const popup = page.waitForEvent("popup");
  await page.evaluate(() => (window as any).__isoler());
  const fen = await popup;
  await fen.locator('html[data-mega-pret="1"]').waitFor({ timeout: 30000 });
  await fen.bringToFront();
  await expect(ao(fen).locator("#ao-run")).toBeVisible();
  await fen.evaluate(() => (window as any).__placerCurseur("a", "fin"));
  await demarrerReleve(fen);
  const trouvees = await laisserTourner(fen, 15000, (a) => a.length >= 3);
  expect(trouvees.length).toBeGreaterThanOrEqual(3);
  expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
  const j = await journal(fen);
  expect(saisiesHors(j, ["a"])).toEqual([]);
  expect(sansInvisibles((await etat(fen)).redaction?.a?.libre)).toBe("");
  expect((await journal(page)).saisies).toEqual([]); // rien dans la fenêtre principale
});
test("RF2 [fr] « Nouveau » ouvre une fenêtre séparée : « Démarrer » depuis la boîte s'arrête sans rien taper", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", fenetre: 1 }));
  const popup = page.waitForEvent("popup");
  await demarrerReleve(page);
  const fen = await popup;
  await fen.locator('html[data-faux-pret="1"]').waitFor();
  await expect.poll(() => noticeReleve(page), { timeout: 30000 }).not.toBe("");
  expect((await journal(page)).saisies).toEqual([]);
  expect((await journal(fen)).saisies).toEqual([]);
  expect(await pastilleReleve(page).isVisible()).toBe(false);
});
test("RF3 [fr] relevé en cours, l'utilisateur clique « Isoler » : arrêt propre, plus aucune frappe dans la fenêtre principale", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", teams: 1 }));
  await messageOuvert(page);
  await demarrerReleve(page);
  await expect(pastilleReleve(page)).toBeVisible({ timeout: 20000 });
  await attendreJusqua(() => adressesReleve(page), (a) => a.length >= 1, 10000, 300);
  const popup = page.waitForEvent("popup");
  await page.evaluate(() => (window as any).__isoler());
  await (await popup).locator('html[data-faux-pret="1"]').waitFor();
  await page.bringToFront();
  await expect(pastilleReleve(page)).toBeHidden({ timeout: 20000 });
  const j = await journal(page);
  expect(saisiesHors(j, ["a"]), "frappe ailleurs après la disparition du champ").toEqual([]);
  expect(await noticeReleve(page)).toMatch(/Arrêté|pause/i);
});

// ============================================================ RG. autres champs destinataires
test("RG1 [it] le champ « A » (italien, une lettre) est bien pris comme champ destinataire", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "it" }));
  await messageOuvert(page);
  await demarrerReleve(page);
  await laisserTourner(page, 8000);
  const j = await journal(page);
  expect(saisiesHors(j, ["a"])).toEqual([]);
  expect(j.saisies.length).toBeGreaterThan(0);
});
test("RG2 [fr] ligne Cci affichée, curseur dans « Cci » : le relevé tape dans « Cci » seulement, sans valider", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", cciVisible: 1 }));
  await messageOuvert(page, { focus: "cci" });
  await demarrerReleve(page);
  const trouvees = await laisserTourner(page, 12000, (a) => a.length >= 3);
  const j = await journal(page);
  expect(saisiesHors(j, ["cci"])).toEqual([]);
  expect(touchesValidation(j)).toEqual([]);
  expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
  const e = await etat(page);
  expect(e.redaction?.cci?.pastilles).toBe(0);
  expect(sansInvisibles(e.redaction?.cci?.libre)).toBe("");
});

// ============================================================ RH. variantes de disposition
test("RH1 [fr] rédaction dans une fenêtre modale (reste de la page aria-hidden) : relève l'annuaire, ne tape que dans « À »", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "fr", dialogue: 1 }));
  await messageOuvert(page);
  await demarrerReleve(page);
  const trouvees = await laisserTourner(page, 15000, (a) => a.length >= 3);
  expect(trouvees.length).toBeGreaterThanOrEqual(3);
  expect(trouvees.filter((x) => !ANNUAIRE.adresses.includes(x))).toEqual([]);
  expect(saisiesHors(await journal(page), ["a"])).toEqual([]);
});
test("RH2 [de] message ouvert, curseur nulle part, « Démarrer » : arrêt sûr, brouillon intact", async ({ context, worker }) => {
  const page = await preparer(context, worker, url({ lang: "de" }));
  await messageOuvert(page, { a: ["alice@exemple.fr"], objet: "Entwurf", corps: "Mein Text", focus: null });
  const avant = await etat(page);
  await demarrerReleve(page);
  await expect.poll(() => noticeReleve(page), { timeout: 20000 }).not.toBe("");
  expect((await journal(page)).saisies).toEqual([]);
  expect((await etat(page)).redaction).toEqual(avant.redaction);
});

// ============================================================ RM. verrou relevé / envoi entre deux fenêtres
test("RM1 [fr] envoi en cours dans la fenêtre principale : « Démarrer » dans la fenêtre isolée refuse de relever", async ({ context, worker }) => {
  // BUG : le verrou mutuel de mega.js ne regarde que la fenêtre courante
  // (PP.moteurActif, adresses/panel.js:448 et 492 ; mega.js:24 et 31) : depuis une fenêtre
  // de rédaction isolée (ou un autre onglet Outlook), le relevé démarre pendant
  // qu'un envoi tourne. Ici il tape dans le brouillon isolé de l'utilisateur (M1).
  const page = await preparer(context, worker, url({ lang: "fr" }));
  await messageOuvert(page, { objet: "Brouillon perso", corps: "À finir", focus: null });
  const popup = page.waitForEvent("popup");
  await page.evaluate(() => (window as any).__isoler());
  const fen = await popup;
  await fen.locator('html[data-mega-pret="1"]').waitFor({ timeout: 30000 });
  // Envoi dans la fenêtre principale, longue pause après le 1er mail.
  await page.bringToFront();
  await ao(page).locator('.mega-tab[data-onglet="envoi"]').click();
  await remplirEnvoi(page, { dest: ["e1@exemple.fr", "e2@exemple.fr"], texte: "Bonjour" });
  await pp(page).locator("#pp-pmin").fill("3000");
  await pp(page).locator("#pp-pmax").fill("3000");
  await pp(page).locator("#pp-go").click();
  await expect.poll(async () => (await envoyesGlobal(page)).length, { timeout: 20000 }).toBe(1);
  // L'utilisateur passe dans la fenêtre isolée et clique « Démarrer ».
  await fen.bringToFront();
  if (!(await ao(fen).isVisible())) await pp(fen).locator('.mega-tab[data-onglet="adresses"]').click();
  await fen.evaluate(() => (window as any).__placerCurseur("a", "fin"));
  await demarrerReleve(fen);
  await fen.waitForTimeout(4000);
  const releveParti = await pastilleReleve(fen).isVisible();
  const j = await journal(fen);
  await pauseReleve(fen);
  expect.soft(j.saisies.length, "frappes du relevé dans la fenêtre isolée pendant l'envoi").toBe(0);
  expect(releveParti, "relevé démarré pendant un envoi").toBe(false);
});
