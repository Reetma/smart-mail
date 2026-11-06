// Audit « arrière-plan et cycle de vie » de Smart Mail : le pilotage d'Outlook web
// doit tenir quand Outlook n'est PAS au premier plan (onglet caché, minuteries
// bridées, gel, rechargement, redémarrage du navigateur, service worker MV3
// arrêté, extension rechargée, deux onglets, veille du Mac), sans aucun dérapage.
//   node node_modules/@playwright/test/cli.js test tests/audit-arriere-plan.spec.ts --workers=2 --reporter=line
// Fausse page : tests/audit/arriere-plan/outlook-reel.html (nouvel Outlook FR,
// structure de la fiche de référence ; l'envoi part en requête réseau enregistrée
// côté Node, donc le journal des envois survit aux rechargements et redémarrages).
// Banc : tests/audit/arriere-plan/banc.ts. Émulation de l'onglet caché :
// tests/audit/arriere-plan/emulation.js (Chromium headless garde toutes les pages
// « visibles » et Playwright coupe le bridage : on reproduit Chrome dans le monde
// de la page ET dans le monde isolé de l'extension ; le gel passe par CDP).
// Mode accéléré (data-pp-test="rapide") : toutes les attentes de l'extension ÷100,
// une « seconde » de pause = 10 ms. Les tests de rechargement à une étape précise
// tournent à vitesse réelle (sinon l'extension irait plus vite que le rechargement).
// Les tests qui échouent aujourd'hui portent un commentaire « // BUG : … ».
import { test as base, expect } from "@playwright/test";
import { Banc, Onglet, URL_OUTLOOK, attendre, campagne, envoisCompte, liste, type CfgPage } from "./audit/arriere-plan/banc";

const test = base.extend<{ ouvrir: (cfg?: CfgPage) => Promise<Banc> }>({
  // eslint-disable-next-line no-empty-pattern
  ouvrir: async ({}, use) => {
    const bancs: Banc[] = [];
    await use(async (cfg: CfgPage = {}) => { const b = await Banc.ouvrir(cfg, 42); bancs.push(b); return b; });
    for (const b of bancs) await b.fin();
  },
});
test.describe.configure({ mode: "parallel" });

const TOUCHES_INTERDITES = ["Enter", "Tab", ";", ","];
const tapees = (o: { page: import("@playwright/test").Page }) => o.page.evaluate(() => ((window as any).__typed as string[]).slice());

// « Reprendre l'envoi ». Si l'extension répond « déjà en cours dans un autre onglet » alors
// qu'aucun autre onglet n'envoie (verrou pp-actif laissé par l'onglet disparu, voir E4),
// on note le refus, on attend l'expiration du verrou (30 s) et on recommence.
async function reprendre(o: Onglet): Promise<boolean> {
  await o.pp().locator("#pp-go").click();
  await attendre(1500);
  if (!((await o.notice().textContent()) ?? "").includes("déjà en cours dans un autre onglet")) return false;
  test.info().annotations.push({ type: "observé", description: "« Reprendre » d'abord refusé : « Un envoi est déjà en cours dans un autre onglet Outlook. » (verrou laissé par l'onglet disparu)" });
  await attendre(31000);
  await o.pp().locator("#pp-go").click();
  return true;
}

// ============================================================ A. envoi, onglet caché
test.describe("A. envoi en onglet caché", () => {
  test("A1 caché juste après « Lancer », puis bridage intensif (1 réveil/min) : un seul mail par destinataire, contenu exact, pauses tenues", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(4, "a1-");
    await o.preparerEnvoi({ dest, objet: "Réunion jeudi", texte: "Bonjour,\nRéunion jeudi à 10 h.\nCordialement", pmin: 300, pmax: 300 }); // 3 s réelles
    await o.lancer();
    await o.cacher("seconde");
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBeGreaterThanOrEqual(1);
    await o.cacher("minute"); // caché depuis plus de 5 min
    await expect.poll(() => b.envois.length, { timeout: 60000 }).toBe(4);
    await attendre(1500);
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    for (const e of b.envois) {
      expect(e.cci).toEqual([]);
      expect(e.objet).toBe("Réunion jeudi");
      expect(e.corps).toContain("Réunion jeudi à 10 h.");
      expect(e.corps.indexOf("Bonjour")).toBeLessThan(e.corps.indexOf("Signature de test"));
    }
    for (let i = 1; i < b.envois.length; i++) {
      const ecart = b.envois[i]!.t - b.envois[i - 1]!.t;
      expect(ecart, `pause entre le mail ${i} et le mail ${i + 1}`).toBeGreaterThanOrEqual(3000 - 100);
      expect(ecart, `pause entre le mail ${i} et le mail ${i + 1} (pas de bridage : 1 réveil/s donnerait ≥ 300 s)`).toBeLessThan(30000);
    }
    await expect(o.notice()).toContainText("Terminé : 4 mail(s) envoyé(s).");
    expect(await envoisCompte(b)).toBe(4);
    expect(b.doublons()).toEqual([]);
  });

  test("A2 caché en plein remplissage (après « À »), puis remis devant pendant la pause : pas de rafale, pause tenue, contenu exact", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "a2-");
    await o.preparerEnvoi({ dest, pmin: 8, pmax: 8 });
    await o.lancer();
    await expect.poll(() => o.contenuA(), { timeout: 30000 }).toContain(dest[0]!);
    await o.cacher("minute");
    await expect.poll(() => b.envois.length, { timeout: 40000 }).toBe(1);
    await attendre(3000);
    await o.montrer(); // l'utilisateur revient pendant la pause
    await expect.poll(() => b.envois.length, { timeout: 90000 }).toBe(3);
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    for (let i = 1; i < 3; i++) expect(b.envois[i]!.t - b.envois[i - 1]!.t).toBeGreaterThanOrEqual(8000);
    expect(b.doublons()).toEqual([]);
  });

  test("A3 variante plausible : la rédaction n'apparaît qu'à l'image suivante (rAF suspendu, onglet caché) → arrêt propre sans envoi, puis rien d'envoyé dans la rédaction restée ouverte", async ({ ouvrir }) => {
    const b = await ouvrir({ redactionRaf: true });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "a3-");
    await o.preparerEnvoi({ dest, pmin: 200, pmax: 200 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await o.cacher("seconde"); // pendant la pause
    await expect(o.notice()).toContainText("Arrêt avant l'envoi à a3-2@exemple.fr (rien n'est parti) : La fenêtre de rédaction ne s'est pas ouverte", { timeout: 30000 });
    expect(b.envois.length).toBe(1);
    await o.montrer(); // l'image en attente est peinte : la rédaction vide apparaît
    await expect.poll(() => o.redactionOuverte()).toBe(true);
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Un message est déjà ouvert en rédaction", { timeout: 15000 });
    expect(b.envois.length).toBe(1);
    await o.page.locator("#abandonner").click();
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Terminé : 2 mail(s) envoyé(s).", { timeout: 30000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("A4 Outlook ne valide le destinataire qu'à sa minuterie (bridée à 1/s, onglet caché, bridage intensif) : destinataire validé, envoi correct", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, pastille: "timeout" });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "a4-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    await o.lancer();
    await o.cacher("minute");
    await expect(o.notice()).toContainText("Terminé : 2 mail(s) envoyé(s).", { timeout: 90000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("A5 Cci en onglet caché : un paquet par mail, « À » vide, personne servi deux fois", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(7, "a5-");
    await o.preparerEnvoi({ dest, pmin: 100, pmax: 100 });
    await o.pp().locator("#pp-cci").check();
    await o.pp().locator("#pp-paquet").fill("3");
    await o.lancer();
    await o.cacher("minute");
    await expect.poll(() => b.envois.length, { timeout: 60000 }).toBe(3);
    await attendre(1000);
    expect(b.envois.map((e) => e.cci)).toEqual([dest.slice(0, 3), dest.slice(3, 6), dest.slice(6)]);
    for (const e of b.envois) expect(e.a).toEqual([]);
    expect(b.doublons()).toEqual([]);
  });
  test("A6 longue pause réelle (40 s) en arrière-plan, bridage intensif : ni raccourcie ni étirée", async ({ ouvrir }) => {
    test.setTimeout(240000);
    const b = await ouvrir({ rapide: false });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "a6-");
    await o.preparerEnvoi({ dest, pmin: 40, pmax: 40 });
    await o.lancer();
    await o.cacher("minute");
    await expect.poll(() => b.envois.length, { timeout: 150000 }).toBe(2);
    const ecart = b.envois[1]!.t - b.envois[0]!.t;
    expect(ecart, "écart (pause 40 s + gestes)").toBeGreaterThanOrEqual(40000);
    expect(ecart, "écart (pas d'étirement à 1 réveil/min)").toBeLessThan(40000 + 25000);
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("A7 importance élevée demandée, onglet caché (bridage intensif) : confirmée avant chaque envoi", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "a7-");
    await o.preparerEnvoi({ dest, pmin: 50, pmax: 50 });
    await o.pp().locator("#pp-importance").check();
    await o.lancer();
    await o.cacher("minute");
    await expect(o.notice()).toContainText("Terminé : 2 mail(s) envoyé(s).", { timeout: 60000 });
    expect(b.envois.map((e) => [e.a[0], e.importance])).toEqual(dest.map((d) => [d, "haute"]));
  });
});

// ============================================================ B. gel / dégel de l'onglet
test.describe("B. onglet gelé puis dégelé", () => {
  test("B1 gelé 20 s pendant la pause : au dégel, pas de rafale, la pause continue, personne servi deux fois", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "b1-");
    await o.preparerEnvoi({ dest, pmin: 600, pmax: 600 }); // 6 s réelles
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    await o.cacher("seconde");
    await o.geler();
    await attendre(20000);
    expect(b.envois.length).toBe(1);
    await o.degeler();
    await attendre(500);
    expect(b.envois.length, "aucun mail parti en rafale au dégel").toBe(1);
    await expect.poll(() => b.envois.length, { timeout: 60000 }).toBe(3);
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    expect(b.envois[2]!.t - b.envois[1]!.t).toBeGreaterThanOrEqual(6000 - 100);
  });

  test("B2 gelé 65 s juste après le clic « Envoyer » (rédaction pas encore refermée) : jamais renvoyé", async ({ ouvrir }) => {
    test.setTimeout(240000);
    const b = await ouvrir({ rapide: false, fermetureLente: 800 });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "b2-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    let gel: Promise<void> | null = null;
    b.surEnvoi = () => { if (b.envois.length === 1) gel = o.geler(); };
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 40000 }).toBe(1);
    await expect.poll(() => gel !== null).toBe(true);
    await gel;
    await attendre(65000);
    await o.degeler();
    await expect(o.notice()).toContainText(/Terminé|à vérifier|ne sera pas renvoyé/, { timeout: 60000 });
    if (!(await o.notice().textContent())?.includes("Terminé")) {
      await expect.poll(() => o.redactionOuverte(), { timeout: 30000 }).toBe(false);
      await o.pp().locator("#pp-go").click();
      await expect(o.notice()).toContainText("Terminé", { timeout: 60000 });
    }
    expect(b.envoisA(dest[0]!).length).toBe(1);
    expect(b.doublons()).toEqual([]);
  });

  test("B3 gelé 20 s en plein remplissage (après « À ») : au dégel, le mail part une seule fois, avec un seul destinataire et le bon contenu", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "b3-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0, objet: "Objet B3", texte: "Texte B3" });
    await o.lancer();
    await expect.poll(() => o.contenuA(), { timeout: 30000 }).toContain(dest[0]!);
    await o.cacher("seconde");
    await o.geler();
    await attendre(20000);
    await o.degeler();
    await o.montrer();
    await expect(o.notice()).toContainText("Terminé : 2 mail(s) envoyé(s).", { timeout: 60000 });
    expect(b.envois.map((e) => [e.a, e.objet, e.corps.includes("Texte B3")])).toEqual(dest.map((d) => [[d], "Objet B3", true]));
  });
});

// ============================================================ C. rechargement à chaque étape d'un envoi
test.describe("C. page rechargée pendant un envoi", () => {
  for (const etape of ["apres-a", "apres-corps"] as const) {
    test(`C1 page rechargée à l'étape « ${etape} » : rien n'est parti, la reprise sert chacun une fois`, async ({ ouvrir }) => {
      const b = await ouvrir({ rapide: false, recharger: etape });
      const o = await b.onglet({ onglet: "envoi" });
      const dest = liste(3, `c1${etape.slice(-1)}-`);
      await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
      await o.lancer();
      await o.apresRechargement();
      await attendre(1500);
      expect(b.envois.length).toBe(0);
      await o.ongletEnvoi();
      await expect(o.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi");
      await expect(o.pp().locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 3");
      await o.pp().locator("#pp-go").click();
      await expect(o.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 90000 });
      expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
      expect(b.doublons()).toEqual([]);
    });
  }

  test("C3 page rechargée au moment de cliquer « Envoyer » : « à vérifier », jamais renvoyé, signalé", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, recharger: "survol-envoyer" });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c3-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    await o.lancer();
    await o.apresRechargement();
    await o.ongletEnvoi();
    await expect(o.pp().locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 3, 1 à vérifier");
    await expect(o.pp().locator("#pp-rows")).toContainText("Page rechargée ou fermée pendant l'envoi");
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Terminé : 2 mail(s) envoyé(s), 1 à vérifier.", { timeout: 90000 });
    expect(b.envois.map((e) => e.a)).toEqual([[dest[1]], [dest[2]]]);
    expect(b.doublons()).toEqual([]);
  });

  test("C3b Cci : page rechargée au moment de cliquer « Envoyer » : tout le paquet « à vérifier », jamais renvoyé", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, recharger: "survol-envoyer" });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(5, "c3b-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    await o.pp().locator("#pp-cci").check();
    await o.pp().locator("#pp-paquet").fill("3");
    await o.lancer();
    await o.apresRechargement();
    await o.ongletEnvoi();
    await expect(o.pp().locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 5, 3 à vérifier");
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Terminé", { timeout: 90000 });
    expect(b.envois.map((e) => e.cci)).toEqual([dest.slice(3)]);
    expect(b.doublons()).toEqual([]);
  });

  test("C4 page rechargée juste après le départ du mail (rédaction pas encore refermée) : la limite sur 24 h doit compter ce mail", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, recharger: "apres-envoi" });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c4-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0, max: 2 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 40000 }).toBe(1);
    await o.apresRechargement();
    await o.ongletEnvoi();
    await expect(o.pp().locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 3, 1 à vérifier");
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText(/Limite de 2 mails sur 24 h atteinte|Terminé/, { timeout: 90000 });
    expect(b.doublons()).toEqual([]);
    // BUG : Moteur.charger() (moteur.js:164-169) change l'« envoi » interrompu en
    // « verifier » sans l'inscrire dans pp-envois, alors que unMail compte toute
    // issue incertaine « par prudence » (moteur.js:316) : 3 mails partent en 24 h
    // pour une limite de 2.
    expect(b.envois.length, `mails partis en 24 h (limite 2), pp-envois = ${await envoisCompte(b)}`).toBeLessThanOrEqual(2);
  });

  test("C5 page rechargée pendant la pause : à la reprise, la pause entre deux mails doit encore être respectée", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c5-");
    await o.preparerEnvoi({ dest, pmin: 2000, pmax: 2000 }); // 20 s réelles entre deux mails
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    await o.recharger(); // Cmd+R, ou Outlook qui se recharge, pendant la pause
    await o.ongletEnvoi();
    await expect(o.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await o.pp().locator("#pp-go").click();
    await expect.poll(() => b.envois.length, { timeout: 60000 }).toBe(2);
    const ecart = b.envois[1]!.t - b.envois[0]!.t;
    expect(b.doublons()).toEqual([]);
    // BUG : lancer() (moteur.js:352-366) part tout de suite sur le mail suivant : ni la
    // campagne ni pp-envois ne servent à imposer le reste de la pause après un
    // rechargement ; deux mails partent à quelques centaines de ms d'intervalle.
    expect(ecart, "écart entre les deux mails (pause réglée : 20 s)").toBeGreaterThanOrEqual(20000 - 500);
  });

  test("C6 Outlook se recharge tout seul pendant la pause, onglet caché : aucune reprise automatique, verrou libéré, la reprise sert chacun une fois", async ({ ouvrir }) => {
    const b = await ouvrir({ autoRecharge: 300 });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c6-");
    await o.preparerEnvoi({ dest, pmin: 1000, pmax: 1000 });
    await o.lancer();
    await o.cacher("seconde");
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await o.apresRechargement();
    await o.cacher("seconde"); // un onglet caché qui se recharge reste caché
    await attendre(15000);
    expect(b.envois.length, "aucun envoi sans l'utilisateur après le rechargement").toBe(1);
    expect((await b.stockage(["pp-actif"]))["pp-actif"], "verrou libéré au déchargement").toBeUndefined();
    expect((await campagne(b))?.curseur).toBe(1);
    await o.montrer();
    await o.ongletEnvoi();
    await expect(o.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 60000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("C7 session expirée pendant la pause (page de connexion) puis retour dans Outlook : rien ne part entre-temps, puis chacun une seule fois", async ({ ouvrir }) => {
    const b = await ouvrir({ redirection: 300 });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c7-");
    await o.preparerEnvoi({ dest, pmin: 1000, pmax: 1000 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.page.url(), { timeout: 30000 }).toContain("login.microsoftonline.com");
    await attendre(12000);
    expect(b.envois.length).toBe(1);
    o.iso = null;
    await o.page.goto(URL_OUTLOOK);
    await o.pret();
    await o.ongletEnvoi();
    await reprendre(o);
    await expect(o.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 60000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("C8 variante plausible : après rechargement Outlook rouvre le brouillon à moitié rempli → la reprise refuse, n'envoie pas ce brouillon et n'y ajoute personne", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, recharger: "apres-a", restaurerBrouillon: true });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c8-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    await o.lancer();
    await o.apresRechargement();
    await expect.poll(() => o.contenuA()).toContain(dest[0]!); // brouillon rouvert par Outlook
    await o.ongletEnvoi();
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Un message est déjà ouvert en rédaction", { timeout: 30000 });
    expect(b.envois.length).toBe(0);
    expect(await o.page.locator("._EType_RECIPIENT_ENTITY").count()).toBe(1);
    await o.page.locator("#abandonner").click();
    await o.pp().locator("#pp-go").click();
    await expect(o.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 90000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });
  test("C9 onglet déchargé (discard) pendant la pause puis rechargé : protégé pendant l'envoi, rien ne part tout seul, reprise sans doublon", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "c9-");
    await o.preparerEnvoi({ dest, pmin: 1500, pmax: 1500 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    const w = await b.sw();
    const avant = await w.evaluate(async () => (await chrome.tabs.query({})).map((t) => ({ id: t.id!, auto: t.autoDiscardable })));
    const cible = avant.find((t) => t.auto === false);
    expect(cible, "onglet Outlook protégé du déchargement automatique pendant l'envoi").toBeTruthy();
    // Déchargé quand même (chrome://discards, autre extension, mémoire critique).
    const nouvelId = await w.evaluate(async (id) => (await chrome.tabs.discard(id))?.id, cible!.id);
    const tDecharge = Date.now();
    await attendre(5000);
    expect(b.envois.length).toBe(1);
    const evtPage = b.context.waitForEvent("page");
    await w.evaluate((id) => chrome.tabs.reload(id), nouvelId!);
    const o2 = new Onglet(b, await evtPage);
    await o2.brancher();
    await o2.pret();
    await o2.ongletEnvoi();
    await expect(o2.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await o2.pp().locator("#pp-go").click();
    await attendre(1500);
    // Verrou laissé par l'onglet déchargé (pas de pagehide) : refus temporaire au plus 30 s.
    if ((await o2.notice().textContent())?.includes("déjà en cours")) {
      test.info().annotations.push({ type: "observé", description: "verrou périmé laissé par l'onglet déchargé : « Reprendre » refusé jusqu'à 30 s" });
      await attendre(Math.max(0, tDecharge + 31000 - Date.now()));
      await o2.pp().locator("#pp-go").click();
    }
    await expect(o2.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 60000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    const apres = await w.evaluate(async () => (await chrome.tabs.query({})).map((t) => t.autoDiscardable));
    expect(apres.every((x) => x !== false), "protection retirée à la fin de l'envoi").toBe(true);
  });
});

// ============================================================ D. extension rechargée, service worker arrêté
test.describe("D. extension rechargée ou service worker arrêté", () => {
  test("D1 service worker arrêté par Chrome à répétition pendant l'envoi (pause, remplissage, clic) : l'envoi continue, une fois chacun, pauses tenues", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(4, "d1-");
    await o.preparerEnvoi({ dest, pmin: 300, pmax: 300 });
    await o.lancer();
    await o.cacher("seconde");
    let arrets = 0, fini = false;
    const boucle = (async () => { while (!fini) { await o.arreterSW().then(() => arrets++, () => {}); await attendre(700); } })();
    await expect.poll(() => b.envois.length, { timeout: 90000 }).toBe(4);
    fini = true; await boucle;
    expect(arrets).toBeGreaterThan(5);
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    for (let i = 1; i < 4; i++) expect(b.envois[i]!.t - b.envois[i - 1]!.t).toBeGreaterThanOrEqual(3000 - 100);
    await expect(o.notice()).toContainText("Terminé : 4 mail(s) envoyé(s).", { timeout: 15000 });
  });

  test("D2 extension rechargée (↻) pendant la pause : le script orphelin ne doit plus rien piloter dans Outlook", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "d2-");
    await o.preparerEnvoi({ dest, pmin: 800, pmax: 800 }); // 8 s réelles
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    await b.rechargerExtension();
    await attendre(15000); // la pause (8 s) est passée
    expect(b.envois.length).toBe(1); // rien n'est parti
    const vu = await o.page.evaluate(() => ({ objet: (document.querySelector('input[aria-label="Objet"]') as HTMLInputElement | null)?.value ?? null, corps: (document.querySelector('[aria-label="Corps du message"]') as HTMLElement | null)?.innerText ?? null }));
    const notice = await o.notice().textContent();
    // BUG : après le rechargement de l'extension, le script de contenu orphelin continue
    // la campagne : tranche() avale l'erreur « Extension context invalidated »
    // (outlook.js:26) et lire() rend undefined (moteur.js:41-43) ; il ouvre un nouveau
    // message et y colle le destinataire suivant, l'objet et le texte, puis s'arrête à
    // avantEnvoi sur « Erreur : Extension context invalidated. ».
    expect(await o.redactionOuverte(), `nouveau message ouvert par le script orphelin (À « ${await o.contenuA()} », objet « ${vu.objet} », notice « ${notice} »)`).toBe(false);
    expect(await o.contenuA(), "destinataire collé par le script orphelin").toBe("");
  });

  test("D3 extension désactivée en plein mail : rien ne part, la campagne reste intacte ; réactivée et page rechargée, chacun servi une fois", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "d3-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    await o.lancer();
    await expect.poll(() => o.contenuA(), { timeout: 30000 }).toContain(dest[0]!);
    await b.activerExtension(false);
    await attendre(12000);
    expect(b.envois.length).toBe(0);
    await expect(o.pp().locator("#pp-perime")).toBeVisible();
    await b.activerExtension(true);
    await o.recharger();
    await o.ongletEnvoi();
    await expect(o.pp().locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 3");
    await reprendre(o);
    await expect(o.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 90000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });
});

// ============================================================ E. fermeture et réouverture du navigateur
test.describe("E. navigateur fermé puis rouvert (même profil)", () => {
  test("E1 fermé pendant la pause, rouvert : rien ne part tout seul, la reprise sert chacun une fois", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "e1-");
    await o.preparerEnvoi({ dest, pmin: 1000, pmax: 1000 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    await b.fermerNavigateur();
    await b.rouvrir();
    const o2 = await b.onglet({});
    await attendre(5000);
    expect(b.envois.length).toBe(1);
    await o2.ongletEnvoi();
    await expect(o2.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await reprendre(o2);
    await expect(o2.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 60000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("E2 fermé juste après le départ d'un mail (rédaction pas encore refermée), rouvert : « à vérifier », jamais renvoyé", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, fermetureLente: 5000 });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "e2-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    let fermeture: Promise<void> | null = null;
    b.surEnvoi = () => { if (b.envois.length === 1 && !fermeture) fermeture = attendre(50).then(() => b.fermerNavigateur()); };
    await o.lancer();
    await expect.poll(() => fermeture !== null, { timeout: 40000 }).toBe(true);
    await fermeture;
    b.surEnvoi = null;
    await b.rouvrir();
    const o2 = await b.onglet({});
    await o2.ongletEnvoi();
    await expect(o2.pp().locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 3, 1 à vérifier");
    await reprendre(o2);
    await expect(o2.notice()).toContainText("Terminé : 2 mail(s) envoyé(s), 1 à vérifier.", { timeout: 90000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    expect(b.doublons()).toEqual([]);
  });

  test("E3 limite sur 24 h à cheval sur deux sessions et un saut d'horloge : 3 mails au plus sur 24 h, reprise après 24 h", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(5, "e3-");
    await o.preparerEnvoi({ dest, pmin: 300, pmax: 300, max: 3 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(2);
    await o.pp().locator("#pp-pause").click();
    await expect(o.notice()).toContainText("Envoi en pause", { timeout: 30000 });
    await b.fermerNavigateur();
    await b.rouvrir();
    const o2 = await b.onglet({});
    await o2.decaler(3600 * 1000); // une heure plus tard
    await o2.ongletEnvoi();
    await o2.pp().locator("#pp-go").click();
    await expect(o2.notice()).toContainText("Limite de 3 mails sur 24 h atteinte", { timeout: 60000 });
    expect(b.envois.length).toBe(3);
    await o2.decaler(25 * 3600 * 1000); // le lendemain
    await o2.pp().locator("#pp-go").click();
    await expect(o2.notice()).toContainText("Terminé", { timeout: 60000 });
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
  });

  test("E4 navigateur fermé pendant la pause puis rouvert : « Reprendre » ne doit pas être refusé par un faux « envoi déjà en cours dans un autre onglet »", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    await o.preparerEnvoi({ dest: liste(2, "e4-"), pmin: 1000, pmax: 1000 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    await b.fermerNavigateur(); // Cmd+Q pendant la pause
    await b.rouvrir();
    const verrou = (await b.stockage(["pp-actif"]))["pp-actif"] as { t: number } | undefined;
    const o2 = await b.onglet({});
    await o2.ongletEnvoi();
    await o2.pp().locator("#pp-go").click();
    await attendre(1500);
    // BUG : le verrou pp-actif n'est libéré qu'au « pagehide », sans attendre l'écriture
    // (moteur.js:273, liberer() non attendu) : à la fermeture du navigateur (comme après
    // une redirection vers la page de connexion, ou une extension rechargée/désactivée,
    // où liberer() échoue en silence, moteur.js:53,266,341) il reste en place, et
    // « Reprendre » est refusé jusqu'à 30 s avec un message faux.
    await expect(o2.notice(), `verrou resté après la fermeture : ${verrou ? Math.round((Date.now() - verrou.t) / 1000) + " s d'âge" : "aucun"}`).not.toContainText("Un envoi est déjà en cours dans un autre onglet Outlook");
  });
});

// ============================================================ F. deux onglets Outlook
test.describe("F. deux onglets Outlook", () => {
  test("F1 onglet A en train d'envoyer (au premier plan) : « Reprendre » dans l'onglet B est refusé avec un message clair", async ({ ouvrir }) => {
    const b = await ouvrir();
    const A = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "f1-");
    await A.preparerEnvoi({ dest, pmin: 1000, pmax: 1000 });
    await A.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    const B = await b.onglet({ nouveau: true });
    await B.ongletEnvoi();
    await B.pp().locator("#pp-go").click();
    await expect(B.notice()).toContainText("Un envoi est déjà en cours dans un autre onglet Outlook.");
    await expect(A.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 90000 });
    expect(b.doublons()).toEqual([]);
  });

  test("F2 onglet A qui envoie en arrière-plan (battement bridé à 1/min) : « Reprendre » dans l'onglet B ne doit jamais lancer un second envoi en parallèle", async ({ ouvrir }) => {
    test.setTimeout(240000);
    const b = await ouvrir({ rapide: false });
    const A = await b.onglet({ onglet: "envoi" });
    const dest = liste(4, "f2-");
    await A.preparerEnvoi({ dest, pmin: 45, pmax: 45 });
    const ongletA = await A.page.evaluate(() => document.documentElement.dataset.onglet);
    await A.lancer();
    await A.cacher("seconde"); // l'utilisateur passe à autre chose
    await expect.poll(() => b.envois.length, { timeout: 40000 }).toBe(1);
    const t1 = b.envois[0]!.t;
    // A est en arrière-plan depuis plus de 5 min (une campagne dure des heures) :
    // Chrome ne réveille plus son setInterval(battement, 5000) qu'une fois par minute.
    await A.cacher("minute");
    // Deuxième onglet Outlook, au premier plan : son panneau propose « Reprendre l'envoi ».
    const B = await b.onglet({ nouveau: true });
    await B.ongletEnvoi();
    await expect(B.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await attendre(Math.max(0, t1 + 33000 - Date.now()));
    await B.pp().locator("#pp-go").click();
    await expect.poll(() => b.envois.filter((e) => e.onglet === ongletA).length, { timeout: 120000 }).toBeGreaterThanOrEqual(2);
    await attendre(1000);
    const parOnglet = b.envois.map((e) => `${e.a[0]} par ${e.onglet === ongletA ? "A" : "B"}`).join(", ");
    // BUG : le verrou entre onglets (pp-actif, moteur.js:258-266) repose sur un battement
    // setInterval de 5 s (moteur.js:334), bridé à 1/min dans un onglet caché ; périmé
    // après 30 s, il laisse B démarrer la même campagne, et A ne relit ni le verrou ni
    // « Déjà envoyés » avant de cliquer : A et B envoient au même destinataire.
    expect(b.doublons(), `destinataires servis deux fois (${parOnglet})`).toEqual([]);
  });

  test("F3 onglet A gelé par Chrome en plein mail (> 30 s) : l'onglet B ne doit pas reprendre la campagne, et A dégelé ne doit pas renvoyer", async ({ ouvrir }) => {
    test.setTimeout(240000);
    const b = await ouvrir({ rapide: false });
    const A = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "f3-");
    await A.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    const ongletA = await A.page.evaluate(() => document.documentElement.dataset.onglet);
    await A.lancer();
    await expect.poll(() => b.envois.length, { timeout: 40000 }).toBe(1);
    const B = await b.onglet({ nouveau: true });
    await B.ongletEnvoi();
    // A remplit le mail suivant ; il passe en arrière-plan et Chrome le gèle.
    await expect.poll(() => A.contenuA(), { timeout: 30000 }).toContain("f3-2@exemple.fr");
    await A.cacher("seconde");
    await A.geler();
    const tGel = Date.now();
    await attendre(36000);
    await B.pp().locator("#pp-go").click();
    await attendre(Math.max(0, tGel + 52000 - Date.now()));
    const pendantGel = b.envois.map((e) => `${e.a[0]} par ${e.onglet === ongletA ? "A" : "B"}`).join(", ");
    await A.degeler();
    await A.montrer();
    await expect.poll(() => b.envois.filter((e) => e.onglet === ongletA).length, { timeout: 60000 }).toBeGreaterThanOrEqual(2);
    await attendre(3000);
    // BUG : même cause que F2 (battement arrêté pendant le gel → verrou périmé) ; A
    // dégelé termine son mail sans revérifier qu'il tient encore le verrou ni que le
    // destinataire a déjà été servi.
    expect(b.doublons(), `destinataires servis deux fois (pendant le gel : ${pendantGel})`).toEqual([]);
  });

  test("F4 campagne terminée dans A, onglet B resté ouvert avec l'ancien état : « Reprendre » dans B n'envoie à personne une seconde fois", async ({ ouvrir }) => {
    const b = await ouvrir();
    const A = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "f4-");
    await A.preparerEnvoi({ dest, pmin: 100, pmax: 100 });
    await A.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    const B = await b.onglet({ nouveau: true }); // B charge la campagne en cours de route
    await B.ongletEnvoi();
    await expect(A.notice()).toContainText("Terminé : 3 mail(s) envoyé(s).", { timeout: 60000 });
    await expect(B.pp().locator("#pp-go")).toHaveText("Reprendre l'envoi"); // état périmé
    await B.pp().locator("#pp-go").click();
    await expect(B.notice()).toContainText("Terminé", { timeout: 30000 });
    expect(b.envois.length).toBe(3);
    expect(b.doublons()).toEqual([]);
  });
});

// ============================================================ G. saut d'horloge (veille du Mac)
test.describe("G. veille du Mac (gel + saut d'horloge)", () => {
  test("G1 veille de 2 h pendant la pause : au réveil, pas de rafale ni de doublon, la pause reprend", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(3, "g1-");
    await o.preparerEnvoi({ dest, pmin: 800, pmax: 800 });
    await o.lancer();
    await expect.poll(() => b.envois.length, { timeout: 30000 }).toBe(1);
    await expect.poll(() => o.enPause(), { timeout: 10000 }).toBe(true);
    await o.cacher("seconde");
    await o.decaler(2 * 3600 * 1000); // l'heure murale saute au moment de la veille…
    await o.geler();                   // …et plus rien ne tourne pendant la veille
    await attendre(3000);
    await o.degeler();
    await attendre(300);
    expect(b.envois.length).toBe(1);
    await expect.poll(() => b.envois.length, { timeout: 60000 }).toBe(3);
    expect(b.envois.map((e) => e.a)).toEqual(dest.map((d) => [d]));
    expect(b.envois[2]!.t - b.envois[1]!.t).toBeGreaterThanOrEqual(8000 - 100);
  });

  test("G2 veille de 2 h juste après le clic « Envoyer » (Outlook n'a pas encore refermé) : le mail n'est jamais renvoyé", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false, fermetureLente: 3000 });
    const o = await b.onglet({ onglet: "envoi" });
    const dest = liste(2, "g2-");
    await o.preparerEnvoi({ dest, pmin: 0, pmax: 0 });
    let saut: Promise<void> | null = null;
    b.surEnvoi = () => { if (b.envois.length === 1) saut = o.decaler(2 * 3600 * 1000); };
    await o.lancer();
    await expect.poll(() => saut !== null, { timeout: 40000 }).toBe(true);
    await saut;
    await expect(o.notice()).toContainText(/Terminé|ne sera pas renvoyé/, { timeout: 60000 });
    if (!(await o.notice().textContent())?.includes("Terminé")) {
      await expect.poll(() => o.redactionOuverte(), { timeout: 30000 }).toBe(false); // l'utilisateur vérifie, la rédaction s'est refermée
      await o.pp().locator("#pp-go").click();
      await expect(o.notice()).toContainText("Terminé", { timeout: 60000 });
    }
    expect(b.envoisA(dest[0]!).length).toBe(1);
    expect(b.envoisA(dest[1]!).length).toBe(1);
  });
});

// ============================================================ H. relevé
async function demarrerReleve(o: Awaited<ReturnType<Banc["onglet"]>>, minimum = 3): Promise<void> {
  await o.ongletAdresses();
  await o.ao().locator("#ao-run").click();
  await expect(o.ao().locator("#ao-stop")).toBeVisible({ timeout: 20000 });
  await expect.poll(async () => (await tapees(o)).length, { timeout: 60000 }).toBeGreaterThanOrEqual(minimum);
}
async function releveSain(o: Awaited<ReturnType<Banc["onglet"]>>): Promise<void> {
  expect(await o.contenuA(), "texte laissé dans « À »").toBe("");
  expect(await o.page.evaluate((t) => ((window as any).__keys as string[]).filter((k) => t.includes(k)), TOUCHES_INTERDITES), "touches de validation").toEqual([]);
  expect(await o.page.evaluate(() => (window as any).__ailleurs), "saisie hors du champ « À »").toEqual([]);
  expect(await o.page.evaluate(() => (window as any).__clicSuggestion ?? 0), "suggestion cliquée").toBe(0);
  expect(await o.page.locator("._EType_RECIPIENT_ENTITY").count(), "destinataire validé").toBe(0);
}

test.describe("H. relevé d'adresses en arrière-plan", () => {
  for (const variante of ["sync", "debounce", "raf-position"] as const) {
    test(`H1 relevé en cours, onglet caché (suggestions « ${variante} ») : arrêt propre à la recherche suivante, message clair, plus aucune frappe, la dernière recherche est juste`, async ({ ouvrir }) => {
      const b = await ouvrir({ suggestions: variante });
      const o = await b.onglet({ onglet: "adresses" });
      await demarrerReleve(o);
      await o.cacher("seconde");
      await expect(o.ao().locator("#ao-notice")).toContainText("n'est plus au premier plan", { timeout: 30000 });
      const n = (await tapees(o)).length;
      await attendre(8000);
      expect((await tapees(o)).length, "frappes après l'arrêt").toBe(n);
      await releveSain(o);
      await expect(o.ao().locator("#ao-run")).toHaveText("Reprendre");
      // Chaque recherche notée au journal (y compris celle finie onglet caché) montre la vraie liste.
      const journal = String((await b.stockage(["reprise-outlook"]))["reprise-outlook"] ?? "");
      const lignes = journal.split("\n").slice(1).filter(Boolean);
      expect(lignes.length).toBeGreaterThan(0);
      for (const l of lignes) {
        const [q, reste] = l.split("\t");
        const vus = (reste ?? "").split(String.fromCharCode(30)).filter(Boolean).map((x) => x.split(String.fromCharCode(31))[1]).sort();
        const attendus = (await o.page.evaluate((x) => (window as any).__ask(x).map((c: { mail: string }) => c.mail), q)).sort();
        expect(vus, `résultat noté pour « ${q} »`).toEqual(attendus);
      }
    });
  }

  test("H2 fenêtre Chrome derrière une autre application (onglet visible, sans le focus) : arrêt propre, rien de tapé ensuite", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "adresses" });
    await demarrerReleve(o);
    await o.sansFocus();
    await expect(o.ao().locator("#ao-notice")).toContainText("n'est plus au premier plan", { timeout: 30000 });
    const n = (await tapees(o)).length;
    await attendre(6000);
    expect((await tapees(o)).length).toBe(n);
    await releveSain(o);
  });

  test("H3 onglet caché pendant l'ouverture du nouveau message par « Démarrer » : rien n'est tapé, message clair", async ({ ouvrir }) => {
    const b = await ouvrir({ rapide: false });
    const o = await b.onglet({ onglet: "adresses" });
    await o.ongletAdresses();
    await o.ao().locator("#ao-run").click();
    await o.cacher("seconde");
    await expect(o.ao().locator("#ao-notice")).toContainText(/premier plan/, { timeout: 30000 });
    await attendre(4000);
    expect(await tapees(o)).toEqual([]);
    expect(await o.page.evaluate(() => (window as any).__segEnCours())).toBe("");
    await releveSain(o);
  });

  test("H4 bridage intensif d'emblée (caché depuis longtemps) : le relevé s'arrête, ne rampe pas et ne boucle pas", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "adresses" });
    await demarrerReleve(o);
    await o.cacher("minute");
    await expect(o.ao().locator("#ao-notice")).toContainText("n'est plus au premier plan", { timeout: 30000 });
    const n = (await tapees(o)).length;
    await attendre(6000);
    expect((await tapees(o)).length).toBe(n);
    await releveSain(o);
  });

  test("H5 Outlook se recharge en plein relevé : rien ne reste tapé, « Reprendre » accepté tout de suite, les recherches faites ne sont pas retapées", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "adresses" });
    await demarrerReleve(o, 5);
    await o.recharger();
    const journal = String((await b.stockage(["reprise-outlook"]))["reprise-outlook"] ?? "");
    const faites = journal.split("\n").slice(1).filter(Boolean).map((l) => l.split("\t")[0]);
    expect(faites.length).toBeGreaterThan(0);
    await o.ongletAdresses();
    await expect(o.ao().locator("#ao-run")).toHaveText("Reprendre");
    await o.ao().locator("#ao-run").click();
    await expect(o.ao().locator("#ao-stop")).toBeVisible({ timeout: 20000 });
    await expect.poll(async () => (await tapees(o)).length, { timeout: 60000 }).toBeGreaterThanOrEqual(3);
    await o.ao().locator("#ao-stop").click();
    await expect(o.ao().locator("#ao-stop")).toBeHidden({ timeout: 20000 });
    const reprises = await tapees(o);
    expect(reprises.filter((q) => faites.includes(q)), "recherches déjà faites, retapées").toEqual([]);
    await releveSain(o);
  });

  test("H6 extension rechargée (↻) pendant le relevé : le relevé doit s'arrêter, et tout ce qu'il affiche comme trouvé doit être enregistré", async ({ ouvrir }) => {
    test.setTimeout(240000);
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "adresses" });
    await demarrerReleve(o);
    const avant = (await tapees(o)).length;
    await b.rechargerExtension();
    await attendre(15000);
    const apres = (await tapees(o)).length;
    if (await o.ao().locator("#ao-stop").isVisible()) await o.ao().locator("#ao-stop").click();
    await expect(o.ao().locator("#ao-stop")).toBeHidden({ timeout: 20000 });
    const affichees = (await o.ao().locator("#ao-rows .m").allTextContents()).map((x) => x.trim().toLowerCase());
    const message = (await o.ao().locator("#ao-notice").textContent()) ?? "";
    expect(await o.contenuA()).toBe("");
    const st = await b.stockage(["base-outlook", "mega-annuaire", "reprise-outlook"]);
    const enregistre = new Set([st["base-outlook"], st["mega-annuaire"], st["reprise-outlook"]].map((x) => String(x ?? "")).join("\n").toLowerCase().match(/[a-z0-9._%+'-]+@[a-z0-9.-]+\.[a-z]{2,}/g) ?? []);
    const perdues = affichees.filter((m) => !enregistre.has(m));
    // BUG : le script orphelin continue de taper dans « À » après le rechargement de
    // l'extension (natif/adapter-io/panel avalent « Extension context invalidated ») ;
    // ses trouvailles ne sont plus enregistrées et le message final affirme pourtant
    // qu'elles sont « dans l'annuaire commun ».
    expect.soft(apres - avant, "recherches tapées par le script orphelin après le rechargement").toBeLessThanOrEqual(2);
    expect(perdues, `adresses affichées comme trouvées mais jamais enregistrées (${perdues.length}/${affichees.length}) ; message : « ${message.slice(0, 140)} »`).toEqual([]);
  });

  test("H7 relevé gelé 20 s en pleine recherche (onglet caché), puis dégelé : il finit la recherche, s'arrête proprement et vide le champ", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "adresses" });
    await demarrerReleve(o);
    await o.cacher("seconde");
    await o.geler();
    await attendre(20000);
    await o.degeler();
    await expect(o.ao().locator("#ao-notice")).toContainText("n'est plus au premier plan", { timeout: 30000 });
    const n = (await tapees(o)).length;
    await attendre(5000);
    expect((await tapees(o)).length).toBe(n);
    await releveSain(o);
  });
  test("H8 navigateur fermé en plein relevé puis rouvert : adresses conservées, et le point de reprise aussi (« Reprendre » sans retaper les recherches faites)", async ({ ouvrir }) => {
    const b = await ouvrir();
    const o = await b.onglet({ onglet: "adresses" });
    await demarrerReleve(o, 6);
    const tapeesAvant = (await tapees(o)).length;
    await b.fermerNavigateur(); // Cmd+Q (ou plantage) en plein relevé
    await b.rouvrir();
    const o2 = await b.onglet({});
    const st = await b.stockage(["reprise-outlook", "mega-annuaire"]);
    const faites = String(st["reprise-outlook"] ?? "").split("\n").slice(1).filter(Boolean).map((l) => l.split("\t")[0]);
    const annuaire = String(st["mega-annuaire"] ?? "").split("\n").filter(Boolean).length;
    expect(annuaire, "adresses relevées conservées dans l'annuaire commun").toBeGreaterThan(0);
    await o2.ongletAdresses();
    // BUG : le journal de reprise n'est écrit qu'en fin de passage : la « minuterie de
    // sécurité » de 2 s est réarmée à chaque ajout (adapter-io.js:80-81), donc ne part
    // jamais tant que les recherches s'enchaînent en moins de 2 s, et le vidage au
    // pagehide (adapter-io.js:98) n'est pas attendu. Fermeture ou plantage du navigateur
    // en plein relevé : le point de reprise est perdu, le panneau propose « Démarrer » et
    // tout est retapé depuis A, sans que rien ne signale l'interruption.
    expect(faites.length, `recherches notées au journal de reprise (tapées avant la fermeture : ${tapeesAvant})`).toBeGreaterThanOrEqual(tapeesAvant - 2);
    await expect(o2.ao().locator("#ao-run")).toHaveText("Reprendre");
  });
});
