// Audit Smart Mail — annuaire commun et listes « Déjà envoyés » sous contrainte (volume, casse, espaces, doublons,
// import/vidage pendant une pause, nouvelles personnes pendant une campagne, adresses mal formées).
// Banc et fausse page : tests/audit/verrous-annuaire/ (banc.ts, outlook.html). Aucune requête réelle.
// Règles visées : E1, E2, E5, M1.
// Lancer : node node_modules/@playwright/test/cli.js test tests/audit-verrous-annuaire-listes.spec.ts --workers=2 --reporter=line
import {
  test, expect, finEnvoi, ao, pp, ouvrirOnglet, montrerOnglet, recharger, preparerEnvoi, lancerEnvoi, pauseEnvoi, attendreEnvois, demarrerReleve,
  attendreReleve, pauseReleve, etat, sonde, journal, lire, poser, liste, destinataires, type Page,
} from "./audit/verrous-annuaire/banc.ts";

async function observerTachesLongues(page: Page) {
  await page.evaluate(() => {
    const w = window as unknown as { __longues: number[] };
    w.__longues = [];
    new PerformanceObserver((l) => l.getEntries().forEach((e) => w.__longues.push(Math.round(e.duration)))).observe({ type: "longtask", buffered: true });
  });
}
async function tachesLongues(page: Page): Promise<number[]> {
  return page.evaluate(() => (window as unknown as { __longues: number[] }).__longues ?? []);
}
// Instant (horloge de la page) du clic sur « Lancer l'envoi », pour mesurer la fenêtre avant le verrou.
async function noterClicLancer(page: Page) {
  await page.evaluate(() => {
    const w = window as unknown as { __clicLancer?: number };
    document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-go")!.addEventListener("click", () => { w.__clicLancer = Date.now(); }, { capture: true });
  });
}
const nombre = (s: string | null) => parseInt(String(s ?? "").replace(/\D/g, "") || "0", 10);
const grosAnnuaire = (n: number) => Array.from({ length: n }, (_, i) => `p${String(i).padStart(5, "0")}@exemple.fr`);

test.describe("Annuaire commun énorme (50 000 adresses)", () => {
  test("VA-L-01 50 000 adresses : chargées et comptées sans gel, envoi aux premières dans l'ordre, liste de l'envoi figée malgré les arrivées", async ({ banc }, info) => {
    const N = 50000, base = grosAnnuaire(N);
    await poser(banc, { "base-outlook": base.join("\n") });
    const t0 = Date.now();
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const tPret = Date.now() - t0;
    await observerTachesLongues(page);
    await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText(/Annuaire commun : 50\s000 personnes relevées/, { timeout: 30000 });
    await expect(pp(page).locator("#pp-nb")).toHaveText(/^50\s000 adresses$/);
    expect(nombre(await ao(page).locator("#mega-send .compte").textContent())).toBe(N);
    await preparerEnvoi(page, { objet: "Cinquante mille", pmin: 0, pmax: 0, max: 3 });
    await noterClicLancer(page);
    await lancerEnvoi(page);
    await finEnvoi(page);
    await expect(pp(page).locator("#pp-notice")).toContainText("Limite de 3 mails sur 24 h atteinte");
    const s = await sonde(page);
    const fenetre = s.debutsEnvoi[0] - (await page.evaluate(() => (window as unknown as { __clicLancer: number }).__clicLancer));
    expect(destinataires(banc)).toEqual(base.slice(0, 3));
    expect((await lire<string[]>(banc, "pp-liste")).length).toBe(N);
    // Arrivées pendant la campagne (relevé ailleurs) : la liste de l'envoi ne bouge pas.
    await poser(banc, { "base-outlook": base.concat(["zz1@exemple.fr", "aa0@exemple.fr"]).join("\n") });
    await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText(/50\s002 personnes relevées/, { timeout: 15000 });
    const l2 = await lire<string[]>(banc, "pp-liste");
    expect(l2.length).toBe(N);
    expect(l2.slice(0, 4)).toEqual(base.slice(0, 4));
    const longues = await tachesLongues(page);
    const mesure = `prêt en ${tPret} ms ; clic → verrou de l'envoi : ${fenetre} ms ; tâches longues (ms) : ${longues.slice(0, 12).join(", ")} (max ${Math.max(0, ...longues)})`;
    info.annotations.push({ type: "mesure", description: mesure });
    console.log("VA-L-01", mesure);
    expect(Math.max(0, ...longues), mesure).toBeLessThan(2000);
  });

  test("VA-L-02 50 000 adresses + relevé en direct : chaque personne trouvée entre dans l'annuaire (enregistré), pas de gel", async ({ banc }, info) => {
    test.setTimeout(150000);
    const N = 50000, base = grosAnnuaire(N);
    await poser(banc, { "base-outlook": base.join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, vue: "adresses" });
    await observerTachesLongues(page);
    await demarrerReleve(page);
    await attendreReleve(page);
    await page.waitForTimeout(15000);
    await pauseReleve(page);
    await page.waitForTimeout(2000); // annuaire enregistré 1 s après la dernière trouvaille
    const trouves = String(await lire<string>(banc, "base-outlook")).split("\n").filter((x) => !x.startsWith("p"));
    const annuaire = String(await lire<string>(banc, "mega-annuaire")).split("\n").filter(Boolean);
    const recherches = (await journal(page)).frappes.filter((f) => !f.confiance && f.zone === "a" && f.key !== "Backspace").length;
    const longues = await tachesLongues(page);
    const mesure = `${trouves.length} trouvées en 15 s, ${recherches} lettres tapées, annuaire enregistré : ${annuaire.length}, tâches longues max ${Math.max(0, ...longues)} ms (${longues.length})`;
    info.annotations.push({ type: "mesure", description: mesure });
    console.log("VA-L-02", mesure);
    expect(trouves.length).toBeGreaterThan(3);
    for (const t of trouves) expect(annuaire).toContain(t);
    expect(annuaire.length).toBe(N + trouves.length);
    expect(nombre(await ao(page).locator("#mega-send .compte").textContent())).toBe(N + trouves.length);
    expect(Math.max(0, ...longues), mesure).toBeLessThan(1500);
  });

  test("VA-L-03 50 000 adresses : « Lancer l'envoi » puis « Démarrer » 300 ms après — le relevé est refusé, jamais les deux à la fois", async ({ banc }, info) => {
    const N = 50000, base = grosAnnuaire(N);
    await poser(banc, { "base-outlook": base.join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", rapide: false, vue: "envoi", delai: 300 });
    await preparerEnvoi(page, { objet: "Course 50k", pmin: 0, pmax: 0, max: 1 });
    await noterClicLancer(page);
    await lancerEnvoi(page);
    await page.waitForTimeout(300);
    await montrerOnglet(page, "adresses");
    await ao(page).locator("#ao-run").click();
    await finEnvoi(page);
    await page.waitForTimeout(1500);
    const s = await sonde(page);
    const fenetre = (s.debutsEnvoi[0] ?? 0) - (await page.evaluate(() => (window as unknown as { __clicLancer: number }).__clicLancer));
    const notice = await ao(page).locator("#ao-notice").textContent();
    await pauseReleve(page);
    info.annotations.push({ type: "mesure", description: `clic → verrou : ${fenetre} ms ; notice relevé : ${notice}` });
    console.log("VA-L-03", fenetre, notice);
    expect(s.ensemble).toBe(0);
    expect(destinataires(banc)).toEqual(base.slice(0, 1));
  });

  test("VA-L-04 50 000 dans l'annuaire, 49 990 « Déjà envoyés » : il reste exactement les 10 autres, envoyées une fois", async ({ banc }) => {
    const N = 50000, base = grosAnnuaire(N);
    const restants = [7, 4242, 9999, 12345, 20000, 33333, 40001, 44444, 49998, 49999].map((i) => base[i]);
    const deja = base.filter((x) => !restants.includes(x)).map((x, i) => (i % 3 === 0 ? x.toUpperCase() : i % 3 === 1 ? ` ${x}\t` : x));
    await poser(banc, { "base-outlook": base.join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await pp(page).locator("#pp-deja-file").setInputFiles({ name: "deja.txt", mimeType: "text/plain", buffer: Buffer.from(deja.join("\r\n")) });
    await expect(pp(page).locator("#pp-notice")).toContainText(/49\s990 adresses ajoutées/, { timeout: 30000 });
    await expect(pp(page).locator("#pp-nb")).toHaveText("10 adresses");
    await preparerEnvoi(page, { objet: "Restants", pmin: 0, pmax: 0 });
    await lancerEnvoi(page);
    await finEnvoi(page);
    expect(destinataires(banc)).toEqual(restants);
  });
});

test.describe("« Déjà envoyés » : casse, espaces, doublons, import et vidage pendant une pause", () => {
  test("VA-L-05 casse, espaces, tabulations, « Nom <adresse> », CSV et doublons : retirés de l'annuaire et des adresses collées", async ({ banc }) => {
    await poser(banc, { "base-outlook": "alice@exemple.fr\nbob@exemple.fr\ncarole@exemple.fr\ndenis@exemple.fr" });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await pp(page).locator("#pp-deja-file").setInputFiles({ name: "deja.csv", mimeType: "text/plain", buffer: Buffer.from('  ALICE@Exemple.FR ;\nBob <BOB@exemple.fr>\t\nbob@exemple.fr\n"carole@exemple.fr";Carole\nemile@EXEMPLE.fr\n') });
    await expect(pp(page).locator("#pp-notice")).toContainText("4 adresses ajoutées");
    await pp(page).locator("#pp-dest").fill("Emile@exemple.fr\n  ALICE@exemple.fr  \nfanny@exemple.fr");
    await expect(pp(page).locator("#pp-nb")).toHaveText("2 adresses");
    await preparerEnvoi(page, { objet: "Casse", pmin: 0, pmax: 0 });
    await lancerEnvoi(page);
    await finEnvoi(page);
    expect(destinataires(banc).sort()).toEqual(["denis@exemple.fr", "fanny@exemple.fr"]);
  });

  test("VA-L-06 « Déjà envoyés » importé pendant une pause (même onglet), envoi à l'annuaire : sauté à la reprise", async ({ banc }) => {
    await poser(banc, { "base-outlook": liste("e", 4).join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { objet: "Pause et import", pmin: 600, pmax: 600 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await pauseEnvoi(page);
    await pp(page).locator("#pp-deja-file").setInputFiles({ name: "deja.txt", mimeType: "text/plain", buffer: Buffer.from("E3@EXEMPLE.FR\n") });
    await expect(pp(page).locator("#pp-deja-nb")).toHaveText("2 adresses");
    await pp(page).locator("#pp-pmin").fill("0");
    await pp(page).locator("#pp-pmax").fill("0");
    await lancerEnvoi(page);
    await finEnvoi(page);
    expect(destinataires(banc)).toEqual(["e1@exemple.fr", "e2@exemple.fr", "e4@exemple.fr"]);
    await expect(pp(page).locator("#pp-suivi-resume")).toContainText("1 déjà servi(s) avant");
  });

  test("VA-L-07 « Déjà envoyés » vidé pendant une pause (même onglet) : la reprise ne renvoie rien de déjà parti", async ({ banc }) => {
    await poser(banc, { "base-outlook": liste("f", 3).join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { objet: "Pause et vidage", pmin: 600, pmax: 600 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await pauseEnvoi(page);
    page.once("dialog", (d) => void d.accept());
    await pp(page).locator("#pp-deja-vider").click();
    await expect(pp(page).locator("#pp-deja-nb")).toHaveText("0 adresse");
    await pp(page).locator("#pp-pmin").fill("0");
    await pp(page).locator("#pp-pmax").fill("0");
    await lancerEnvoi(page);
    await finEnvoi(page);
    expect(destinataires(banc)).toEqual(liste("f", 3));
  });

  test("VA-L-08 Cci par paquets à l'annuaire, « Déjà envoyés » importés pendant la pause : jamais dans un paquet, jamais deux fois", async ({ banc }) => {
    const L = liste("h", 10);
    await poser(banc, { "base-outlook": L.join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { objet: "Cci", pmin: 600, pmax: 600, max: 1000 });
    await pp(page).locator("#pp-cci").check();
    await pp(page).locator("#pp-paquet").fill("3");
    await page.waitForTimeout(400);
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await pauseEnvoi(page);
    await pp(page).locator("#pp-deja-file").setInputFiles({ name: "deja.txt", mimeType: "text/plain", buffer: Buffer.from("h5@exemple.fr\nh8@exemple.fr\n") });
    await pp(page).locator("#pp-pmin").fill("0");
    await pp(page).locator("#pp-pmax").fill("0");
    await lancerEnvoi(page);
    await finEnvoi(page);
    const tous = banc.envois.flatMap((m) => m.cci);
    expect(banc.envois.every((m) => m.a.length === 0)).toBe(true);
    expect(tous.sort()).toEqual(L.filter((x) => x !== "h5@exemple.fr" && x !== "h8@exemple.fr").sort());
  });

  test("VA-L-09 arrivées puis « Vider l'historique » d'Adresses pendant une campagne à l'annuaire : la liste de l'envoi ne bouge pas", async ({ banc }) => {
    await poser(banc, { "base-outlook": liste("k", 3).join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { objet: "Figée", pmin: 300, pmax: 300 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await poser(banc, { "base-outlook": liste("k", 6).join("\n") });
    await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText("Annuaire commun : 6 personnes relevées");
    await banc.sw.evaluate(() => chrome.storage.local.remove(["base-outlook", "dernier-outlook"])); // ⚙ Adresses > Vider l'historique
    await finEnvoi(page);
    expect(destinataires(banc)).toEqual(liste("k", 3));
    expect(await lire<string[]>(banc, "pp-liste")).toEqual(liste("k", 3));
  });

  test("VA-L-10 case « Annuaire commun » décochée puis onglet rechargé pendant une campagne : la campagne reprend sur SA liste", async ({ banc }) => {
    await poser(banc, { "base-outlook": liste("q", 3).join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(page, { objet: "Décochée", pmin: 600, pmax: 600 });
    await lancerEnvoi(page);
    await attendreEnvois(banc, 1);
    await pauseEnvoi(page);
    await poser(banc, { "mega-annuaire-actif": false });
    await recharger(page);
    await expect(pp(page).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await pp(page).locator("#pp-pmin").fill("0");
    await pp(page).locator("#pp-pmax").fill("0");
    await lancerEnvoi(page);
    await finEnvoi(page);
    expect(destinataires(banc)).toEqual(liste("q", 3));
  });
});

test.describe("Adresses de l'annuaire et des listes : jamais une autre adresse que celle prévue", () => {
  test("VA-L-11 adresse importée par « Adresses » avec apostrophe de tête ou domaine invalide : jamais tapée telle quelle dans « À »", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "adresses" });
    await ao(page).locator("#ao-gear").click();
    await ao(page).locator("#ao-file").setInputFiles({ name: "export.txt", mimeType: "text/plain", buffer: Buffer.from("'bob@exemple.fr'\ncarole@exemple..fr\ndenis@exemple.fr\n") });
    await expect(ao(page).locator("#ao-notice")).toContainText("3 adresse(s) ajoutée(s) à l'historique");
    await ao(page).locator("#ao-gear").click();
    await ao(page).locator("#mega-send").click();
    await expect(pp(page)).toBeVisible();
    // bob a déjà reçu le message (liste « Déjà envoyés » d'un précédent envoi)
    await pp(page).locator("#pp-deja-file").setInputFiles({ name: "deja.txt", mimeType: "text/plain", buffer: Buffer.from("bob@exemple.fr\n") });
    await expect(pp(page).locator("#pp-deja-nb")).toHaveText("1 adresse");
    const nb = await pp(page).locator("#pp-nb").textContent();
    await preparerEnvoi(page, { objet: "Annuaire importé", pmin: 0, pmax: 0 });
    await lancerEnvoi(page);
    await finEnvoi(page);
    const notice = await pp(page).locator("#pp-notice").textContent();
    // BUG : l'annuaire (mega.js ajouter) n'est ni nettoyé ni validé comme les adresses collées (Moteur.extraireAdresses +
    // O.MAIL) : « 'bob@exemple.fr » échappe à « Déjà envoyés » et part, « carole@exemple..fr » est collée dans « À »
    // et l'envoi s'arrête « à vérifier » après avoir cliqué « Envoyer ».
    expect(destinataires(banc), `liste proposée : ${nb} ; fin : ${notice}`).toEqual(["denis@exemple.fr"]);
  });

  test("VA-L-12 adresse avec accent ou « & » dans la liste collée : jamais tronquée en une autre adresse", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const saisies = ["hélène.dupont@exemple.fr", "r&d@exemple.fr", "zoe.fabre@exemple.fr"];
    await preparerEnvoi(page, { dest: saisies, objet: "Tronquées", pmin: 0, pmax: 0 });
    const nb = await pp(page).locator("#pp-nb").textContent();
    await lancerEnvoi(page);
    await finEnvoi(page);
    const exactes = saisies.map((x) => x.toLowerCase());
    // BUG : Moteur.extraireAdresses (regex TROUVE sans ancrage sur les bords du mot) extrait « ne.dupont@exemple.fr »
    // de « hélène.dupont@… » et « d@exemple.fr » de « r&d@… » : le mail part à une AUTRE adresse, sans avertissement.
    expect(destinataires(banc).filter((x) => !exactes.includes(x.toLowerCase())), `liste comptée : ${nb}`).toEqual([]);
  });

  test("VA-L-13 adresse avec accent importée par « Adresses » puis envoyée via l'annuaire : jamais tronquée", async ({ banc }) => {
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "adresses" });
    await ao(page).locator("#ao-gear").click();
    await ao(page).locator("#ao-file").setInputFiles({ name: "liste.txt", mimeType: "text/plain", buffer: Buffer.from("hélène.dupont@exemple.fr\nzoe.fabre@exemple.fr\n") });
    await expect(ao(page).locator("#ao-notice")).toContainText("ajoutée(s) à l'historique");
    const base = String(await lire<string>(banc, "base-outlook")).split("\n").filter(Boolean);
    await ao(page).locator("#ao-gear").click();
    await ao(page).locator("#mega-send").click();
    await preparerEnvoi(page, { objet: "Accents", pmin: 0, pmax: 0 });
    await lancerEnvoi(page);
    await finEnvoi(page);
    // BUG : même extraction tronquée côté « Adresses » (panel.js importerTexte, regex EMAIL) : « ne.dupont@exemple.fr »
    // entre dans l'annuaire commun et reçoit le mail.
    expect(destinataires(banc).filter((x) => !["hélène.dupont@exemple.fr", "zoe.fabre@exemple.fr"].includes(x)), `historique : ${base.join(", ")}`).toEqual([]);
  });
});

test.describe("Annuaire : mêmes personnes vues par plusieurs sources", () => {
  test("VA-L-14 même personne relevée sur Outlook, sur Gmail et collée (casses différentes) : servie une seule fois", async ({ banc }) => {
    await poser(banc, { "base-outlook": "Jean.Dupont@Exemple.fr\nalice@exemple.fr", "base-gmail": "jean.dupont@exemple.fr\nALICE@EXEMPLE.FR\nbob@exemple.fr", "mega-annuaire": "JEAN.DUPONT@EXEMPLE.FR\n" });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText("Annuaire commun : 3 personnes relevées");
    await pp(page).locator("#pp-dest").fill("Jean.DUPONT@exemple.fr\nBob@Exemple.FR");
    await expect(pp(page).locator("#pp-nb")).toHaveText("3 adresses");
    await preparerEnvoi(page, { objet: "Sources", pmin: 0, pmax: 0 });
    await lancerEnvoi(page);
    await finEnvoi(page);
    expect(destinataires(banc).map((x) => x.toLowerCase()).sort()).toEqual(["alice@exemple.fr", "bob@exemple.fr", "jean.dupont@exemple.fr"]);
  });

  test("VA-L-15 « Envoyer à l'annuaire » répété entre deux relevés : chaque nouvel envoi ne sert que les nouvelles personnes", async ({ banc }) => {
    await poser(banc, { "base-outlook": liste("j", 2).join("\n") });
    const page = await ouvrirOnglet(banc, { onglet: "A", vue: "adresses" });
    for (const [n, objet] of [[2, "Vague 1"], [4, "Vague 2"], [5, "Vague 3"]] as Array<[number, string]>) {
      await poser(banc, { "base-outlook": liste("j", n).join("\n") }); // un relevé (ailleurs) a complété l'annuaire
      await expect(ao(page).locator("#mega-send .compte")).toHaveText(String(n));
      await ao(page).locator("#mega-send").click();
      await ao(page).locator("#mega-send").click({ timeout: 800 }).catch(() => {});  // 2e clic (panneau déjà basculé) : sans effet
      if (await pp(page).locator("#pp-abandon").isVisible()) await pp(page).locator("#pp-abandon").click(); // « Nouvel envoi »
      await preparerEnvoi(page, { objet, pmin: 0, pmax: 0 });
      await lancerEnvoi(page);
      await finEnvoi(page);
      await montrerOnglet(page, "adresses");
    }
    expect(destinataires(banc)).toEqual(liste("j", 5));
    expect(banc.envois.map((m) => m.objet)).toEqual(["Vague 1", "Vague 1", "Vague 2", "Vague 2", "Vague 3"]);
  });
});
