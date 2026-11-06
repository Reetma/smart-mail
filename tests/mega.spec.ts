// Bout en bout : Smart Mail (extension/) chargée dans Chromium, contre les
// fausses pages des deux extensions d'origine. /mail/ = fausse page du
// Publipostage (elle enregistre les mails « envoyés »), /releve/ = fausse page
// d'Adresses (champ « À » qui propose des contacts). Aucune requête ne sort.
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
const EXT = join(ICI, "..", "extension");
const PAGE_ENVOI = readFileSync(join(ICI, "fausse-outlook.html"), "utf8");
const PAGE_RELEVE = readFileSync(join(ICI, "fausse-adresses", "outlook.html"), "utf8");
const BOOK = readFileSync(join(ICI, "fausse-adresses", "book.js"), "utf8");

const test = base.extend<{ context: BrowserContext; worker: Worker }>({
  // eslint-disable-next-line no-empty-pattern
  context: async ({}, use) => {
    const dir = mkdtempSync(join(tmpdir(), "mega-outlook-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    await context.route("https://outlook.office.com/**", (route) => {
      const p = new URL(route.request().url()).pathname;
      if (p.endsWith("/book.js")) return route.fulfill({ status: 200, contentType: "text/javascript; charset=utf-8", body: BOOK });
      if (p.startsWith("/releve/")) return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: PAGE_RELEVE });
      return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: PAGE_ENVOI });
    });
    await use(context);
    await context.close();
  },
  worker: async ({ context }, use) => {
    let [w] = context.serviceWorkers();
    if (!w) w = await context.waitForEvent("serviceworker");
    await use(w);
  },
});

const ao = (page: Page) => page.locator('[data-testid="ao-panel"]');
const pp = (page: Page) => page.locator("[data-pp-panneau]");
const onglet = (page: Page, panneau: "ao" | "pp", o: "adresses" | "envoi") => (panneau === "ao" ? ao(page) : pp(page)).locator(`.mega-tab[data-onglet="${o}"]`);
const envoyes = async (page: Page) => JSON.parse((await page.locator("html").getAttribute("data-envoyes")) ?? "[]") as Array<{ a: string[]; objet: string }>;

async function ouvrir(page: Page, url: string) {
  await page.goto(url);
  await page.bringToFront();
  await page.locator('html[data-mega-pret="1"]').waitFor({ timeout: 30000 });
}

test("un seul panneau à deux onglets, au même endroit, onglet retenu au rechargement", async ({ context }) => {
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/mail/");
  await expect(ao(page)).toBeVisible();
  await expect(pp(page)).toBeHidden();
  await expect(onglet(page, "ao", "adresses")).toHaveAttribute("aria-selected", "true");

  await onglet(page, "ao", "envoi").click();
  await expect(pp(page)).toBeVisible();
  await expect(ao(page)).toBeHidden();
  const a = await pp(page).boundingBox();
  expect(a!.x + a!.width).toBeGreaterThan(page.viewportSize()!.width - 30); // collé à droite, plus décalé à gauche d'Adresses

  // Déplacé à la main côté Envoyer : Relever s'ouvre au même endroit.
  const tete = pp(page).locator("#pp-head h1");
  const b = await tete.boundingBox();
  await page.mouse.move(b!.x + 5, b!.y + 5); await page.mouse.down();
  await page.mouse.move(b!.x - 195, b!.y + 105, { steps: 5 }); await page.mouse.up();
  const avant = await pp(page).boundingBox();
  await onglet(page, "pp", "adresses").click();
  await expect(ao(page)).toBeVisible();
  const apres = await ao(page).boundingBox();
  expect(Math.abs(apres!.x - avant!.x)).toBeLessThan(2);
  expect(Math.abs(apres!.y - avant!.y)).toBeLessThan(2);

  await onglet(page, "ao", "envoi").click();
  await page.reload();
  await page.locator('html[data-mega-pret="1"]').waitFor();
  await expect(pp(page)).toBeVisible();
  await expect(ao(page)).toBeHidden();

  // La croix d'Adresses masque tout (ne détruit plus le panneau).
  await onglet(page, "pp", "adresses").click();
  await ao(page).locator("#ao-close").click();
  await expect(ao(page)).toBeHidden();
  await expect(pp(page)).toBeHidden();
  await expect(page.locator('[data-testid="ao-panel"]')).toHaveCount(1);
});

const nbAdresses = async (page: Page) => parseInt(((await pp(page).locator("#pp-nb").textContent()) ?? "0").replace(/\D/g, "") || "0", 10);

test("annuaire commun : la liste d'envoi se remplit EN DIRECT pendant le relevé", async ({ context, worker }) => {
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/releve/outlook.html");
  await page.locator('input[aria-label="To"]').click();
  await ao(page).locator("#ao-run").click(); // curseur dans « À » : on démarre là
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeVisible({ timeout: 20000 }); // relevé en cours
  await onglet(page, "ao", "envoi").click(); // on regarde l'onglet Envoyer pendant que le relevé tourne
  await expect(pp(page).locator("#pp-annuaire")).toBeVisible();
  await expect.poll(() => nbAdresses(page), { timeout: 30000 }).toBeGreaterThanOrEqual(3);
  const avant = await nbAdresses(page);
  await expect.poll(() => nbAdresses(page), { timeout: 30000 }).toBeGreaterThan(avant); // ça monte tout seul
  await expect(pp(page).locator("#pp-annuaire-texte")).toContainText("Annuaire commun : ");
  await expect(onglet(page, "pp", "adresses").locator(".pastille")).toBeVisible(); // toujours en cours

  await onglet(page, "pp", "adresses").click();
  await ao(page).locator("#ao-stop").click();
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeHidden({ timeout: 15000 });
  await expect(ao(page).locator("#mega-send")).toBeEnabled();
  const tout = await worker.evaluate(() => chrome.storage.local.get("base-outlook"));
  const releve = String(tout["base-outlook"] || "").split("\n").filter(Boolean);
  await expect(ao(page).locator("#mega-send .compte")).toHaveText(String(releve.length));
  await ao(page).locator("#mega-send").click();
  await expect(pp(page)).toBeVisible();
  await expect(pp(page).locator("#pp-nb")).toHaveText(`${releve.length} adresses`);

  // Rechargé : l'annuaire est toujours là.
  await page.reload();
  await page.locator('html[data-mega-pret="1"]').waitFor();
  await expect(pp(page).locator("#pp-nb")).toHaveText(`${releve.length} adresses`);
});

test("annuaire : sans doublon avec le texte collé, case décochée retenue, vidé avec l'historique", async ({ context, worker }) => {
  await worker.evaluate(() => chrome.storage.local.set({ "base-outlook": "alice@exemple.fr\nbob@exemple.fr", "base-gmail": "carole@exemple.org" }));
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/mail/");
  await onglet(page, "ao", "envoi").click();
  await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText("Annuaire commun : 3 personnes relevées"); // Outlook + Gmail
  await expect(pp(page).locator("#pp-nb")).toHaveText("3 adresses");
  await pp(page).locator("#pp-dest").fill("Bob@exemple.fr\ndenis@exemple.fr");
  await expect(pp(page).locator("#pp-nb")).toHaveText("4 adresses");

  // Un relevé fait ailleurs (autre onglet) complète l'annuaire.
  await worker.evaluate(() => chrome.storage.local.set({ "base-outlook": "alice@exemple.fr\nbob@exemple.fr\nemma@exemple.fr" }));
  await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText("Annuaire commun : 4 personnes relevées");
  await expect(pp(page).locator("#pp-nb")).toHaveText("5 adresses");

  await pp(page).locator("#pp-annuaire-actif").uncheck();
  await expect(pp(page).locator("#pp-nb")).toHaveText("2 adresses"); // seulement le texte collé
  await page.reload();
  await page.locator('html[data-mega-pret="1"]').waitFor();
  await expect(pp(page).locator("#pp-annuaire-actif")).not.toBeChecked();
  await expect(pp(page).locator("#pp-nb")).toHaveText("2 adresses");
  // « Envoyer à l'annuaire » depuis Relever recoche la case.
  await onglet(page, "pp", "adresses").click();
  await ao(page).locator("#mega-send").click();
  await expect(pp(page).locator("#pp-annuaire-actif")).toBeChecked();
  await expect(pp(page).locator("#pp-nb")).toHaveText("5 adresses");

  // ⚙ Adresses > Vider l'historique (Outlook) : il reste la personne relevée sur Gmail.
  await onglet(page, "pp", "adresses").click();
  page.once("dialog", (d) => d.accept());
  await ao(page).locator("#ao-gear").click();
  await ao(page).locator("#ao-clear-base").click();
  await onglet(page, "ao", "envoi").click();
  await expect(pp(page).locator("#pp-annuaire-texte")).toHaveText("Annuaire commun : 1 personne relevée");
});

test("pendant un relevé, l'envoi refuse de démarrer", async ({ context }) => {
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/releve/outlook.html");
  await page.locator('input[aria-label="To"]').click();
  await ao(page).locator("#ao-run").click(); // curseur dans « À » : on démarre là
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeVisible({ timeout: 20000 });
  await onglet(page, "ao", "envoi").click();
  await pp(page).locator("#pp-dest").fill("alice@exemple.fr");
  await pp(page).locator("#pp-go").click();
  await expect(pp(page).locator("#pp-notice")).toContainText("Un relevé d'adresses est en cours");
  await expect(pp(page).locator("#pp-live")).toBeHidden();
  await onglet(page, "pp", "adresses").click();
  await ao(page).locator("#ao-stop").click();
});

test("envoi à l'annuaire ; pendant l'envoi le relevé ne démarre ni seul ni au bouton, et la liste de l'envoi ne bouge pas", async ({ context, worker }) => {
  await worker.evaluate(() => chrome.storage.local.set({ "base-outlook": "alice@exemple.fr\nbob@exemple.fr\ncarole@exemple.org" }));
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/mail/");
  await expect(ao(page).locator("#ao-distinct")).toHaveText("3 adresses distinctes");
  await ao(page).locator("#mega-send").click();
  await expect(pp(page).locator("#pp-nb")).toHaveText("3 adresses");
  await pp(page).locator("#pp-objet").fill("Invitation");
  await pp(page).locator("#pp-texte").fill("Bonjour,\n\nVoici l'invitation.");
  // Pause longue entre deux mails : on a le temps d'essayer de lancer un relevé.
  await pp(page).locator("#pp-pmin").fill("2000"); // ≈ 20 s en mode test (fausse page)
  await pp(page).locator("#pp-pmax").fill("2000");
  await pp(page).locator("#pp-go").click();

  await expect.poll(async () => (await envoyes(page)).length, { timeout: 15000 }).toBe(1);
  await expect(onglet(page, "pp", "envoi").locator(".pastille")).toBeVisible();
  // Une nouvelle personne arrive dans l'annuaire pendant l'envoi : pas dans CET envoi.
  await worker.evaluate(() => chrome.storage.local.set({ "base-outlook": "alice@exemple.fr\nbob@exemple.fr\ncarole@exemple.org\ndenis@exemple.fr" }));
  await onglet(page, "pp", "adresses").click();
  await expect(ao(page).locator("#mega-send .compte")).toHaveText("4");
  await ao(page).locator("#ao-run").click();
  await expect(ao(page).locator("#ao-notice")).toContainText("Un envoi Publipostage est en cours");
  await expect(ao(page).locator("#ao-stop")).toBeHidden();
  expect(await envoyes(page)).toHaveLength(1); // toujours dans la pause

  await expect.poll(async () => (await envoyes(page)).length, { timeout: 90000 }).toBe(3);
  expect((await envoyes(page)).map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"], ["carole@exemple.org"]]);
  await expect(page.locator('[data-testid="ao-panel"][data-ao-armed="1"]')).toHaveCount(0);
  await expect(onglet(page, "ao", "envoi").locator(".pastille")).toBeHidden({ timeout: 10000 });

  // Nouvel envoi : seule la nouvelle personne reste (les 3 autres sont « Déjà envoyés »).
  await onglet(page, "ao", "envoi").click();
  await pp(page).locator("#pp-abandon").click();
  await expect(pp(page).locator("#pp-nb")).toHaveText("1 adresse");
});

test("plus de démarrage automatique : cliquer et taper dans « À » ne lance rien, « Démarrer » lance", async ({ context }) => {
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/releve/outlook.html");
  await expect(ao(page).locator("#ao-run")).toHaveText("Démarrer");
  await expect(ao(page).locator("#ao-arm-msg")).toContainText("depuis n'importe quelle page d'Outlook");
  await expect(ao(page).locator(".bar:visible")).toHaveCount(0); // plus de barre de progression
  // Bouton principal lisible au survol (le style de survol générique effaçait le dégradé).
  await ao(page).locator("#ao-run").hover();
  await page.waitForTimeout(250);
  expect(await ao(page).locator("#ao-run").evaluate((b) => getComputedStyle(b).backgroundImage)).toContain("gradient");
  await onglet(page, "ao", "envoi").click();
  await pp(page).locator("#pp-go").hover();
  await page.waitForTimeout(250);
  expect(await pp(page).locator("#pp-go").evaluate((b) => getComputedStyle(b).backgroundImage)).toContain("gradient");
  await onglet(page, "pp", "adresses").click();
  await page.locator('input[aria-label="To"]').click();
  await page.locator('input[aria-label="To"]').pressSequentially("a", { delay: 5 });
  await page.waitForTimeout(4000);
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeHidden();
  expect(await page.evaluate(() => (window as any).__typed?.length ?? 0)).toBe(0);
  // Le champ contient le « a » de l'utilisateur : Démarrer REFUSE et n'y touche pas.
  await ao(page).locator("#ao-run").click();
  await expect(ao(page).locator("#ao-notice")).toContainText("je n'y touche pas");
  await expect(page.locator('input[aria-label="To"]')).toHaveValue("a");
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeHidden();
  // Champ vidé par l'utilisateur : Démarrer lance.
  await page.locator('input[aria-label="To"]').fill("");
  await page.locator('input[aria-label="To"]').click();
  await ao(page).locator("#ao-run").click();
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeVisible({ timeout: 15000 });
  await expect(ao(page).locator("#ao-stop")).toHaveText("Pause");
  await expect(ao(page).locator(".bar:visible")).toHaveCount(0);
  await expect.poll(async () => parseInt((await ao(page).locator("#ao-live-eta").textContent())?.match(/(\d+) adresses/)?.[1] ?? "0", 10), { timeout: 60000 }).toBeGreaterThanOrEqual(5);
  await ao(page).locator("#ao-stop").click();
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeHidden({ timeout: 15000 });

  // En pause : « Reprendre » continue là où ça s'est arrêté (rien n'est retapé).
  await expect(ao(page).locator("#ao-resume")).toBeVisible();
  await expect(ao(page).locator("#ao-resume")).toContainText("En pause");
  await expect(ao(page).locator("#ao-notice")).toContainText("En pause. Tout ce qui a été trouvé est dans l'annuaire commun.");
  await expect(ao(page).locator("#ao-export")).toBeEnabled();
  await expect(ao(page).locator("#ao-run")).toHaveText("Reprendre");
  const tapees = await page.evaluate(() => [...((window as any).__typed as string[])]);
  await page.locator('input[aria-label="To"]').click();
  await ao(page).locator("#ao-run").click();
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeVisible({ timeout: 15000 });
  await expect.poll(async () => (await page.evaluate(() => (window as any).__typed.length)), { timeout: 30000 }).toBeGreaterThan(tapees.length + 1);
  const suite = await page.evaluate((n) => ((window as any).__typed as string[]).slice(n), tapees.length);
  // la reprise ne recommence pas au début : aucune des recherches déjà faites n'est retapée
  // (la recherche abandonnée par la pause, elle, est refaite)
  expect(suite.filter((q) => tapees.slice(0, -1).includes(q))).toEqual([]);
  await ao(page).locator("#ao-stop").click();
});

test("« Démarrer » depuis la boîte de réception : ouvre un nouveau message, puis pause, export possible, reprise", async ({ context, worker }) => {
  await worker.evaluate(() => chrome.storage.local.set({ "base-outlook": "alice@exemple.fr\nbob@exemple.fr" }));
  const page = context.pages()[0] ?? (await context.newPage());
  await ouvrir(page, "https://outlook.office.com/mail/");
  await expect(page.locator('input[aria-label="Ajouter un objet"]')).toHaveCount(0); // aucun message ouvert
  await ao(page).locator("#ao-run").click();
  await expect(page.locator('input[aria-label="Ajouter un objet"]')).toBeVisible({ timeout: 15000 }); // nouveau message ouvert par l'extension
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeVisible({ timeout: 15000 });
  await expect(ao(page).locator("#ao-export")).toBeDisabled(); // pas pendant le relevé

  await ao(page).locator("#ao-stop").click();
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeHidden({ timeout: 15000 });
  await expect(ao(page).locator("#ao-notice")).toContainText("En pause");
  // En pause : on peut exporter (sans en avoir besoin) et reprendre.
  await expect(ao(page).locator("#ao-export")).toBeEnabled();
  await expect(ao(page).locator("#ao-copy")).toBeEnabled();
  const [dl] = await Promise.all([page.waitForEvent("download"), ao(page).locator("#ao-export").click()]);
  expect(dl.suggestedFilename()).toMatch(/^adresses-outlook-.*\.txt$/);
  // Redémarrer : le message est toujours ouvert, on y retourne sans en rouvrir un autre.
  await ao(page).locator("#ao-run").click();
  await expect(onglet(page, "ao", "adresses").locator(".pastille")).toBeVisible({ timeout: 15000 });
  await expect(page.locator('input[aria-label="Ajouter un objet"]')).toHaveCount(1);
  await ao(page).locator("#ao-stop").click();
});
