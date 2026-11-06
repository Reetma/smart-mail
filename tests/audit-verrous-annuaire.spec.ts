// Audit Smart Mail — « Verrous entre relevé et envoi » (même onglet Outlook).
// Banc et fausse page : tests/audit/verrous-annuaire/ (banc.ts, outlook.html). Aucune requête réelle.
// Règles visées : M1 (verrou mutuel, la frappe de l'un ne déclenche jamais l'autre), E2, E3, E5, R4, R5, R6.
// Lancer : node node_modules/@playwright/test/cli.js test tests/audit-verrous-annuaire.spec.ts --workers=2 --reporter=line
//
// Repères :
//  - « sonde » : MutationObserver sur la vue « Relevé en cours » et la vue d'envoi : ensemble > 0 = les deux
//    moteurs ont tourné en même temps.
//  - journal de la fausse page : le relevé (et son diagnostic) tape avec des keydown SYNTHÉTIQUES (isTrusted=false) ;
//    le Publipostage n'émet jamais de keydown (collage + execCommand). Un keydown synthétique dans l'objet, le corps
//    ou un champ destinataire d'un mail en cours d'envoi = frappe du relevé dans ce mail.
//  - pages « rapide » : attentes du Publipostage divisées par 100 (data-pp-test) ; le relevé garde son rythme réel.
import {
  test, expect, finEnvoi, ao, pp, ouvrirOnglet, montrerOnglet, preparerEnvoi, lancerEnvoi, attendreEnvois, demarrerReleve, attendreReleve,
  pauseReleve, etat, sonde, journal, cfgPage, poser, boutonExtension, badges, lire, liste, pauseEnvoi, destinataires, type Page, type Banc,
} from "./audit/verrous-annuaire/banc.ts";

const ZONES_MAIL = ["a", "cc", "cci", "objet", "corps"];
async function frappesExtension(page: Page, zones: string[], depuis = 0) {
  return (await journal(page)).frappes.filter((f) => !f.confiance && zones.includes(f.zone) && f.t >= depuis);
}
async function clicsExtension(page: Page, depuis = 0) {
  return (await journal(page)).clics.filter((c) => !c.confiance && c.t >= depuis);
}
async function texteChamp(page: Page, zone: string): Promise<string> {
  return page.evaluate((z) => { const c = document.querySelector(`[data-champ="${z}"]`) as HTMLElement | null; return c ? ((c as HTMLInputElement).value ?? c.innerText) : "(absent)"; }, zone);
}
async function pastillesA(page: Page): Promise<string[]> {
  return page.evaluate(() => [...document.querySelectorAll('[data-champ="a"] ._EType_RECIPIENT_ENTITY')].map((p) => (p as HTMLElement).dataset.adresse ?? ""));
}
async function ouvrirMessageUtilisateur(page: Page) {
  await page.locator("#nouveau").click();
  await page.locator('[data-champ="a"]').waitFor();
}
// Clic programmé (même tâche) sur des boutons de panneau, même masqués : impossible à la souris dans l'interface à
// onglets, sert à éprouver le verrou lui-même « à quelques millisecondes ».
async function clicsMemeTache(page: Page, ordre: Array<"releve" | "envoi">) {
  await page.evaluate((o) => {
    for (const x of o) {
      const host = document.querySelector(x === "releve" ? '[data-testid="ao-panel"]' : "[data-pp-panneau]") as HTMLElement;
      (host.shadowRoot!.getElementById(x === "releve" ? "ao-run" : "pp-go") as HTMLButtonElement).click();
    }
  }, ordre);
}

test.describe("M1 — clics rapprochés sur « Démarrer » et « Lancer l'envoi » (même onglet)", () => {
  test("VA-M1-01 « Démarrer » (aucun message ouvert) puis « Lancer l'envoi » 400 ms après : jamais les deux à la fois, mail exact", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 900, vue: "envoi" });
    await preparerEnvoi(page, { dest: ["course1@exemple.fr"], objet: "Course 1", texte: "Bonjour,\nCourse un." });
    await montrerOnglet(page, "adresses");
    const t0 = Date.now();
    await ao(page).locator("#ao-run").click(); // le relevé ouvre un nouveau message (≈ 1 à 2 s)
    await page.waitForTimeout(150);
    await montrerOnglet(page, "envoi");
    await lancerEnvoi(page);
    await finEnvoi(page);
    await page.waitForTimeout(2500);
    await pauseReleve(page);
    const s = await sonde(page);
    expect(s.ensemble, "relevé et envoi en cours en même temps").toBe(0);
    for (const m of banc.envois) { expect(m.a).toEqual(["course1@exemple.fr"]); expect(m.cc).toEqual([]); expect(m.objet).toBe("Course 1"); }
    expect(await frappesExtension(page, ["objet", "corps", "cc", "cci"], t0)).toEqual([]);
  });

  test("VA-M1-02 « Lancer l'envoi » puis « Démarrer » 300 ms après : le relevé refuse, rien n'est tapé", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 300, vue: "envoi" });
    await preparerEnvoi(page, { dest: ["course2@exemple.fr"], objet: "Course 2" });
    const t0 = Date.now();
    await lancerEnvoi(page);
    await page.waitForTimeout(150);
    await montrerOnglet(page, "adresses");
    await ao(page).locator("#ao-run").click();
    await expect(ao(page).locator("#ao-notice")).toContainText("Un envoi Publipostage est en cours");
    await finEnvoi(page);
    const s = await sonde(page);
    expect(s.releve).toBe(0);
    expect(s.debutsReleve).toEqual([]);
    expect(banc.envois.map((m) => m.a)).toEqual([["course2@exemple.fr"]]);
    expect(await frappesExtension(page, ZONES_MAIL, t0)).toEqual([]);
  });

  for (const ordre of [["releve", "envoi"], ["envoi", "releve"]] as Array<Array<"releve" | "envoi">>) {
    test(`VA-M1-03 clics dans la même tâche (${ordre.join(" puis ")}), aucun message ouvert : un seul moteur, mail exact`, async ({ banc }) => {
      const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 700, vue: "envoi" });
      await preparerEnvoi(page, { dest: ["meme1@exemple.fr"], objet: "Même tâche" });
      const t0 = Date.now();
      await clicsMemeTache(page, ordre);
      await finEnvoi(page);
      await page.waitForTimeout(2000);
      const s = await sonde(page);
      await pauseReleve(page);
      expect(s.ensemble, "relevé et envoi en cours en même temps").toBe(0);
      for (const m of banc.envois) expect(m.a).toEqual(["meme1@exemple.fr"]);
      expect(await frappesExtension(page, ["objet", "corps", "cc", "cci"], t0)).toEqual([]);
    });

    test(`VA-M1-04 clics dans la même tâche (${ordre.join(" puis ")}), curseur dans « À » d'un message ouvert : verrou tenu`, async ({ banc }) => {
      const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 300, vue: "envoi" });
      await preparerEnvoi(page, { dest: ["meme2@exemple.fr"], objet: "Même tâche 2" });
      await ouvrirMessageUtilisateur(page);
      await page.locator('[data-champ="a"]').click();
      const t0 = Date.now();
      await clicsMemeTache(page, ordre);
      await finEnvoi(page);
      await page.waitForTimeout(1500);
      const s = await sonde(page);
      await pauseReleve(page);
      // Le Publipostage finit par refuser (« Un message est déjà ouvert en rédaction ») : rien n'est collé dans le
      // message ouvert. Mais le verrou lui-même n'est pas atomique : les deux moteurs démarrent.
      const e = await etat(page);
      const fr = await frappesExtension(page, ["a"], t0);
      expect(await pastillesA(page)).toEqual([]);
      expect(await texteChamp(page, "objet")).toBe("");
      // BUG : verrou « vérifier puis agir » avec des attentes entre les deux (Publipostage : moteur.creer puis
      // actifAilleurs ; relevé : runActiveElsewhere) : les deux moteurs passent le contrôle et démarrent ensemble.
      expect(s.ensemble, `relevé et envoi en cours en même temps (envoi : « ${e.ppNotice} » ; relevé : ${fr.length} frappes dans « À », ${s.debutsReleve.length} démarrage(s))`).toBe(0);
      void banc;
    });
  }

  test("VA-M1-05 « Démarrer » 5 fois pendant la pause entre deux mails : toujours refusé, l'envoi continue intact", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { dest: ["p1@exemple.fr", "p2@exemple.fr"], objet: "Pause", pmin: 400, pmax: 400 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await montrerOnglet(page, "adresses");
    for (let i = 0; i < 5; i++) await ao(page).locator("#ao-run").click();
    await expect(ao(page).locator("#ao-notice")).toContainText("Un envoi Publipostage est en cours");
    await attendreEnvois(banc, 2);
    await finEnvoi(page);
    const s = await sonde(page);
    expect(s.releve).toBe(0);
    expect(banc.envois.map((m) => m.a[0])).toEqual(["p1@exemple.fr", "p2@exemple.fr"]);
  });

  test("VA-M1-06 « Démarrer » juste après la fin d'un envoi : le relevé part dans un NOUVEAU message, pas dans un reste", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi", delai: 60 });
    await preparerEnvoi(page, { dest: ["fin1@exemple.fr"], objet: "Fin" });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await finEnvoi(page);
    await demarrerReleve(page);
    await attendreReleve(page);
    await expect.poll(async () => (await frappesExtension(page, ["a"])).length, { timeout: 30000 }).toBeGreaterThan(1);
    await pauseReleve(page);
    expect(await pastillesA(page)).toEqual([]);
    expect(await texteChamp(page, "objet")).toBe("");
    expect((await journal(page)).redactions).toBe(2);
    expect(banc.envois).toHaveLength(1);
  });
});

test.describe("M1 — la frappe de l'un ne déclenche jamais l'autre", () => {
  test("VA-M1-07 la frappe du Publipostage dans « À » (insertText, événements isTrusted) n'arme ni ne lance le relevé", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await cfgPage(page, { aCollerIgnore: true }); // Outlook ignore le collage : l'extension insère elle-même (execCommand)
    await preparerEnvoi(page, { dest: ["martin.dupont@exemple.fr", "marie.martin@exemple.fr", "emma.girard@exemple.fr"], objet: "Frappe", pmin: 50, pmax: 50 });
    await lancerEnvoi(page);
    await page.waitForTimeout(300);
    await montrerOnglet(page, "adresses"); // l'utilisateur regarde l'onglet Relever pendant l'envoi
    await attendreEnvois(banc, 3);
    await finEnvoi(page);
    await page.waitForTimeout(3000); // laisse passer un éventuel armement différé (2 s)
    const s = await sonde(page);
    const saisiesA = (await journal(page)).saisies.filter((x) => x.zone === "a" && x.confiance);
    expect(saisiesA.length, "le Publipostage a bien produit des saisies isTrusted dans « À »").toBeGreaterThan(0);
    expect(s.arme, "panneau Adresses armé par la frappe du Publipostage").toBe(0);
    expect(s.debutsReleve).toEqual([]);
    expect(s.releve).toBe(0);
    expect(banc.envois.map((m) => m.a)).toEqual([["martin.dupont@exemple.fr"], ["marie.martin@exemple.fr"], ["emma.girard@exemple.fr"]]);
  });

  test("VA-M1-08 diagnostic d'« Adresses » lancé pendant un envoi : rien n'est tapé dans le mail en cours", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 300, vue: "envoi" });
    await preparerEnvoi(page, { dest: ["diag1@exemple.fr"], objet: "Objet exact", texte: "Bonjour,\nLigne deux du message." });
    const t0 = Date.now();
    await lancerEnvoi(page);
    // Le Publipostage vient de coller le message dans le corps (le curseur y est) ; l'utilisateur, qui attend,
    // ouvre ⚙ d'« Adresses » et lance le diagnostic (« Rien n'est envoyé »).
    await page.locator("html[data-corps-colle]").waitFor({ state: "attached", timeout: 60000 });
    await montrerOnglet(page, "adresses");
    await ao(page).locator("#ao-gear").click();
    await ao(page).locator("#ao-diag").click();
    await finEnvoi(page);
    const intrus = await frappesExtension(page, ZONES_MAIL, t0);
    const corps = banc.envois[0]?.corps ?? "";
    // BUG : le diagnostic n'est pas couvert par le verrou (panel.js : ao-diag.onclick ne teste pas PP.moteurActif,
    // armHere ignore la suspension) : il arme le champ où est le curseur — le corps du mail en cours d'envoi —,
    // y tape « a » puis efface 3 caractères ; le mail part modifié (ou l'envoi s'arrête).
    expect(intrus.map((f) => `${f.zone}:${f.key}`), "frappes du relevé dans le mail en cours d'envoi").toEqual([]);
    expect(banc.envois).toHaveLength(1);
    expect(corps).toMatch(/Ligne deux du message\.\s*\n/);
    expect(corps).not.toMatch(/message\.a/);
  });

  test("VA-M1-09 diagnostic du Publipostage lancé pendant un relevé : refusé, aucun clic dans la rédaction du relevé", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 300, vue: "adresses" });
    await demarrerReleve(page); // ouvre un nouveau message et tape dans « À »
    await attendreReleve(page);
    await expect.poll(async () => (await frappesExtension(page, ["a"])).length, { timeout: 30000 }).toBeGreaterThan(2);
    const t1 = Date.now();
    await montrerOnglet(page, "envoi");
    await pp(page).locator("#pp-gear").click();
    await pp(page).locator("#pp-diag").click();
    await page.waitForTimeout(8000);
    const clics = (await clicsExtension(page, t1)).filter((c) => !/plus de résultats/i.test(c.cible)).map((c) => c.cible);
    const imp = await page.locator("#imp-haute").getAttribute("aria-pressed").catch(() => null);
    const notice2 = await pp(page).locator("#pp-notice2").textContent();
    await pauseReleve(page);
    // BUG : pp-diag.onclick ne teste que moteur.enCours (pas AOPanneau.enCours()) : O.diagnostic() prend la rédaction
    // du relevé, clique « Cci », les onglets du ruban et « Importance élevée » pendant que le relevé y tape.
    expect(clics, `clics du diagnostic pendant le relevé (notice : ${notice2})`).toEqual([]);
    expect(imp).not.toBe("true");
  });
});

test.describe("Changement d'onglet du panneau, masquage, « Envoyer à l'annuaire »", () => {
  test("VA-M1-10 changer d'onglet pendant le démarrage du relevé : il démarre proprement, ne tape que dans « À », le vide à la pause", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 1200, vue: "adresses" });
    await ao(page).locator("#ao-run").click();
    await montrerOnglet(page, "envoi");
    await page.waitForTimeout(300);
    await montrerOnglet(page, "adresses");
    await attendreReleve(page, 30000);
    await expect.poll(async () => (await frappesExtension(page, ["a"])).length, { timeout: 30000 }).toBeGreaterThan(3);
    await pauseReleve(page);
    expect(await frappesExtension(page, ["objet", "corps", "cc", "cci", "recherche", "liste"])).toEqual([]);
    expect((await texteChamp(page, "a")).replace(/[​\s]/g, "")).toBe("");
    expect((await sonde(page)).ensemble).toBe(0);
  });

  test("VA-M1-11 panneau masqué par le bouton de l'extension pendant un envoi : l'envoi continue, rien n'est perdu ni doublé", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { dest: ["m1@exemple.fr", "m2@exemple.fr", "m3@exemple.fr"], objet: "Masqué", pmin: 100, pmax: 100 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await boutonExtension(banc, page);
    await expect(pp(page)).toBeHidden();
    await expect(ao(page)).toBeHidden();
    await attendreEnvois(banc, 3);
    await page.waitForTimeout(1500);
    await boutonExtension(banc, page);
    await expect(pp(page)).toBeVisible();
    await expect(pp(page).locator("#pp-notice")).toContainText("Terminé : 3 mail(s) envoyé(s).");
    expect(banc.envois.map((m) => m.a[0])).toEqual(["m1@exemple.fr", "m2@exemple.fr", "m3@exemple.fr"]);
    expect((await sonde(page)).releve).toBe(0);
  });

  test("VA-M1-12 panneau masqué pendant un envoi qui s'arrête sur une anomalie : l'arrêt est signalé", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { dest: ["v1@exemple.fr", "v2@exemple.fr"], objet: "Masqué anomalie", pmin: 100, pmax: 100 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await boutonExtension(banc, page); // masqué
    await cfgPage(page, { envoiBloque: true }); // 2e mail : « Envoyer » ne referme pas la rédaction
    await expect.poll(async () => (await etat(page)).ppNotice, { timeout: 30000 }).toContain("Ce destinataire ne sera pas renvoyé");
    await page.waitForTimeout(1000);
    const badge = (await badges(banc)).join(",");
    const visible = await pp(page).isVisible();
    // BUG : masqué, l'envoi s'arrête (« à vérifier ») sans rien montrer : le panneau reste caché, aucun badge sur
    // l'icône. L'utilisateur croit que l'envoi continue (mega.js : masque n'est levé que par le bouton).
    expect(visible || badge !== "", `arrêt signalé (panneau visible=${visible}, badge="${badge}")`).toBe(true);
  });

  test("VA-M1-13 « Envoyer à l'annuaire » double-cliqué : le 2e clic ne déclenche rien dans l'onglet Envoyer", async ({ banc }) => {
    await poser(banc, { "base-outlook": "alice@exemple.fr\nbob@exemple.fr\ncarole@exemple.fr" });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "adresses" });
    let fichiers = 0;
    page.on("filechooser", () => { fichiers++; });
    await expect(ao(page).locator("#mega-send")).toBeEnabled();
    const b = (await ao(page).locator("#mega-send").boundingBox())!;
    await ao(page).locator("#mega-send").dblclick();
    await expect(pp(page)).toBeVisible();
    const sous = await page.evaluate(([x, y]) => { const h = document.elementFromPoint(x, y) as HTMLElement; const e = h?.shadowRoot?.elementFromPoint(x, y) as HTMLElement | null; return e ? `${e.tagName.toLowerCase()}#${e.id}` : h?.tagName; }, [b.x + b.width / 2, b.y + b.height / 2]);
    await page.waitForTimeout(800);
    expect(fichiers, `2e clic tombé sur ${sous}`).toBe(0);
    await expect(pp(page).locator("#pp-annuaire-actif")).toBeChecked();
    await expect(pp(page).locator("#pp-nb")).toHaveText("3 adresses");
    expect((await sonde(page)).envoi).toBe(0);
    expect(banc.envois).toEqual([]);
  });

  test("VA-M1-14 « Envoyer à l'annuaire » cliqué 5 fois (case décochée avant) : la case revient, aucun envoi ne part", async ({ banc }) => {
    await poser(banc, { "base-outlook": "alice@exemple.fr\nbob@exemple.fr", "mega-annuaire-actif": false });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await expect(pp(page).locator("#pp-annuaire-actif")).not.toBeChecked();
    await montrerOnglet(page, "adresses");
    for (let i = 0; i < 5; i++) {
      await ao(page).locator("#mega-send").click();
      await expect(pp(page)).toBeVisible();
      await montrerOnglet(page, "adresses");
    }
    await montrerOnglet(page, "envoi");
    await expect(pp(page).locator("#pp-annuaire-actif")).toBeChecked();
    await expect(pp(page).locator("#pp-nb")).toHaveText("2 adresses");
    await page.waitForTimeout(1000);
    expect((await sonde(page)).envoi).toBe(0);
    expect(banc.envois).toEqual([]);
  });
});

test.describe("Relevé et envoi l'un après l'autre : restes de l'autre moteur", () => {
  test("VA-M1-15 relevé mis en pause (brouillon resté ouvert) puis « Lancer l'envoi » : refus propre, rien n'est collé dans ce brouillon", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "adresses", delai: 60 });
    await demarrerReleve(page);
    await attendreReleve(page);
    await expect.poll(async () => (await frappesExtension(page, ["a"])).length, { timeout: 30000 }).toBeGreaterThan(2);
    await pauseReleve(page);
    expect((await texteChamp(page, "a")).replace(/[​\s]/g, "")).toBe("");
    await montrerOnglet(page, "envoi");
    await preparerEnvoi(page, { dest: ["apres1@exemple.fr"], objet: "Après relevé" });
    await lancerEnvoi(page);
    await finEnvoi(page);
    await expect(pp(page).locator("#pp-notice")).toContainText("Un message est déjà ouvert en rédaction");
    expect(await pastillesA(page)).toEqual([]);
    expect(await texteChamp(page, "objet")).toBe("");
    expect(banc.envois).toEqual([]);
  });

  test("VA-M1-16 envoi arrêté avant « Envoyer » (curseur resté dans le corps) puis « Démarrer » : rien n'est tapé ni effacé dans ce corps", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi", delai: 60 });
    await cfgPage(page, { collageEnTrop: true }); // garde-fou du Publipostage : « plus de texte que le message prévu »
    await preparerEnvoi(page, { dest: ["arret1@exemple.fr"], objet: "Arrêt", texte: "Bonjour,\nCe message doit rester intact." });
    await lancerEnvoi(page);
    await finEnvoi(page);
    await expect(pp(page).locator("#pp-notice")).toContainText("plus de texte que le message prévu");
    const corpsAvant = await texteChamp(page, "corps");
    const t1 = Date.now();
    await montrerOnglet(page, "adresses");
    await ao(page).locator("#ao-run").click();
    await page.waitForTimeout(6000);
    const dansCorps = await frappesExtension(page, ["corps"], t1);
    const corpsApres = await texteChamp(page, "corps");
    await pauseReleve(page);
    // BUG : champActif() (panel.js) accepte tout contenteditable sauf objet/recherche : le corps du mail laissé
    // ouvert par l'envoi arrêté est « armé », le relevé y tape ses recherches et son 1er effacement (jusqu'à 200
    // retours arrière) mange le texte du message.
    expect.soft(dansCorps.map((f) => (f.key === "Backspace" ? "⌫" : f.key)).join(""), "frappes du relevé dans le corps").toBe("");
    expect(corpsApres, "corps du mail laissé ouvert par l'envoi").toBe(corpsAvant);
    expect(banc.envois).toEqual([]);
  });

  test("VA-M1-17 envoi « à vérifier » (rédaction restée ouverte) puis « Démarrer » : la pastille du destinataire reste intacte", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi", delai: 60 });
    await cfgPage(page, { envoiBloque: true });
    await preparerEnvoi(page, { dest: ["verif1@exemple.fr"], objet: "À vérifier" });
    await lancerEnvoi(page);
    await finEnvoi(page);
    await expect(pp(page).locator("#pp-notice")).toContainText("Ce destinataire ne sera pas renvoyé");
    expect(await pastillesA(page)).toEqual(["verif1@exemple.fr"]);
    await montrerOnglet(page, "adresses");
    await ao(page).locator("#ao-run").click();
    await page.waitForTimeout(6000);
    const pastilles = await pastillesA(page);
    await pauseReleve(page);
    expect(pastilles).toEqual(["verif1@exemple.fr"]);
    expect(await pastillesA(page)).toEqual(["verif1@exemple.fr"]);
    expect(await texteChamp(page, "objet")).toBe("À vérifier");
  });
});

test.describe("Masquer le panneau pendant un relevé, reprendre l'envoi pendant un relevé", () => {
  test("VA-M1-18 pendant un relevé, « Reprendre l'envoi » d'une campagne en pause est refusé et ne touche pas à la campagne enregistrée", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi", delai: 60 });
    await preparerEnvoi(page, { dest: liste("u", 3), objet: "En pause", pmin: 600, pmax: 600 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await pauseEnvoi(page);
    const avant = JSON.stringify(await lire(banc, "pp-campagne"));
    await demarrerReleve(page);
    await attendreReleve(page);
    await lancerEnvoi(page);
    await expect(pp(page).locator("#pp-notice")).toContainText("Un relevé d'adresses est en cours");
    expect(JSON.stringify(await lire(banc, "pp-campagne"))).toBe(avant);
    await pauseReleve(page);
    expect(banc.envois).toHaveLength(1);
    expect((await sonde(page)).ensemble).toBe(0);
  });

  for (const moyen of ["croix du panneau", "bouton de l'extension"]) {
    test(`VA-M1-19 relevé en cours, panneau masqué (${moyen}) : plus aucune frappe sans panneau ni bouton « Pause » visibles`, async ({ banc }) => {
      const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 300, vue: "adresses" });
      await demarrerReleve(page);
      await attendreReleve(page);
      await expect.poll(async () => (await frappesExtension(page, ["a"])).length, { timeout: 30000 }).toBeGreaterThan(2);
      if (moyen === "croix du panneau") await ao(page).locator("#ao-close").click();
      else await boutonExtension(banc, page);
      await expect(ao(page)).toBeHidden();
      await expect(pp(page)).toBeHidden();
      const t1 = Date.now();
      await page.waitForTimeout(7000);
      const apres = await frappesExtension(page, ["a"], t1 + 1500);
      const enCours = (await etat(page)).releve;
      await boutonExtension(banc, page); // réafficher pour remettre en pause
      await pauseReleve(page);
      // BUG : mega.js détourne la croix d'« Adresses » (qui arrêtait le passage : « sinon un passage partirait sans
      // panneau pour l'arrêter ») vers basculer(), qui ne fait que masquer : le relevé continue de taper dans « À »
      // (et d'y ramener le curseur à chaque recherche) sans panneau visible ni bouton « Pause ».
      expect(apres.map((f) => f.key).join(""), `relevé toujours en cours après masquage : ${enCours}`).toBe("");
    });
  }
});

test.describe("Doubles clics sur les boutons de démarrage", () => {
  test("VA-M1-20 double-clic sur « Lancer l'envoi » puis sur « Reprendre l'envoi » : un seul moteur, aucun doublon", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { dest: liste("dc", 3), objet: "Double clic", pmin: 300, pmax: 300 });
    await pp(page).locator("#pp-go").dblclick();
    await attendreEnvois(banc, 1);
    await pauseEnvoi(page);
    await pp(page).locator("#pp-pmin").fill("0");
    await pp(page).locator("#pp-pmax").fill("0");
    await pp(page).locator("#pp-go").dblclick();
    await expect.poll(() => banc.envois.length, { timeout: 30000 }).toBe(3);
    await page.waitForTimeout(1500);
    expect(destinataires(banc)).toEqual(liste("dc", 3));
    expect((await sonde(page)).envoi).toBe(2);
  });

  test("VA-M1-21 double-clic sur « Démarrer » (aucun message ouvert) : un seul nouveau message, un seul relevé", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 600, vue: "adresses" });
    await ao(page).locator("#ao-run").dblclick();
    await attendreReleve(page);
    await expect.poll(async () => (await frappesExtension(page, ["a"])).length, { timeout: 30000 }).toBeGreaterThan(2);
    await pauseReleve(page);
    expect((await journal(page)).redactions).toBe(1);
    expect((await sonde(page)).debutsReleve).toHaveLength(1);
  });
});
