// Bout en bout de l'application Smart Mail (PWA) contre le serveur lancé en mode
// démonstration (SMARTMAIL_DEMO=1, SMARTMAIL_FAUX_CLAUDE=1) :
//   APP_URL=http://127.0.0.1:8766 CAPTURES=/chemin node node_modules/@playwright/test/cli.js test tests/app.spec.ts
import { test, expect, devices, type Page } from "@playwright/test";

const URL_APP = process.env.APP_URL ?? "http://127.0.0.1:8766";
const CAPTURES = process.env.CAPTURES ?? "";

// Format iPhone (taille, écran Retina, tactile, mobile) sans changer de navigateur.
const { defaultBrowserType: _ignore, ...iphone } = devices["iPhone 15"];
const formats = [
  { nom: "iphone", options: iphone },
  { nom: "ordinateur", options: { viewport: { width: 1280, height: 860 }, deviceScaleFactor: 2 } },
];

async function capture(page: Page, format: string, nom: string) {
  if (CAPTURES) { await page.waitForTimeout(450); await page.screenshot({ path: `${CAPTURES}/${format}-${nom}.png` }); } // après l'animation d'entrée
}
const email = (nom: string) => `${nom}-${Date.now()}-${Math.floor(Math.random() * 1e6)}@demo-ecole.fr`;
const nombre = async (page: Page, sel: string) => Number((await page.locator(sel).textContent())?.replace(/\D/g, "") || 0);

async function connecter(page: Page, adresse: string) {
  await page.goto(URL_APP);
  await expect(page.locator("#e-connexion")).toBeVisible();
  await page.locator("#email-demo").fill(adresse);
  await page.locator('#form-demo button[type="submit"]').click();
  await expect(page.locator("#e-objectif")).toBeVisible();
}

for (const f of formats) {
  test.describe(f.nom, () => {
    test.use(f.options as any);

    test(`parcours complet (${f.nom})`, async ({ page }) => {
      const erreurs: string[] = [];
      page.on("pageerror", (e) => erreurs.push(String(e)));
      page.on("console", (m) => { if (m.type() === "error") erreurs.push(m.text()); });
      const adresse = email("parcours");

      await page.goto(URL_APP);
      await expect(page.locator("#e-connexion")).toBeVisible();
      await expect(page.locator("#badge-demo")).toBeVisible();
      await expect(page.locator("#connexion-microsoft")).toBeHidden(); // démo : connexion par e-mail
      await page.locator("#email-demo").fill(adresse);
      await capture(page, f.nom, "1-connexion");
      await page.locator('#form-demo button[type="submit"]').click();

      await expect(page.locator("#e-objectif")).toBeVisible();
      await expect(page.locator("#bonjour")).toContainText(adresse);
      await expect(page.locator("#reprise")).toBeHidden(); // compte neuf
      await expect(page.locator('.etapes-nav li[data-n="1"]')).toHaveClass(/active/);
      await page.locator('#raccourcis .puce[data-v="100"]').click();
      await expect(page.locator("#objectif")).toHaveValue("100");
      await capture(page, f.nom, "2-objectif");

      await page.locator("#start").click();
      await expect(page.locator("#e-lancement")).toBeVisible();
      await expect(page.locator('#etapes li[data-e="2"]')).toHaveClass(/encours/, { timeout: 5000 });
      await capture(page, f.nom, "3-lancement");
      await expect(page.locator("#liste li")).toHaveCount(0); // rien n'est montré pendant l'analyse

      await expect(page.locator("#e-recolte")).toBeVisible({ timeout: 20000 });
      await expect(page.locator("#logique-resume")).toContainText("prenom.nom@demo-ecole.fr");
      await expect(page.locator("#logique-resume")).not.toContainText("Logique trouvée");
      await expect.poll(() => nombre(page, "#trouves"), { timeout: 10000 }).toBeGreaterThan(5);
      await capture(page, f.nom, "4-recolte-en-cours");
      await expect(page.locator("#trouves")).toHaveText("100", { timeout: 30000 });
      await expect(page.locator("#statut-texte")).toHaveText("Objectif atteint");
      await capture(page, f.nom, "5-objectif-atteint");

      const adresses = await page.locator("#liste li .a").allTextContents();
      expect(adresses).toHaveLength(100);
      for (const a of adresses) expect(a, a).toMatch(/^[a-z-]{2,}\.[a-z-]+\d*@demo-ecole\.fr$/);
      expect(await nombre(page, "#ecartes")).toBeGreaterThan(0);

      await page.locator("#vers-envoi").click();
      await expect(page.locator("#e-envoi")).toBeVisible();
      await expect(page.locator("#r-dest")).toHaveText("100 étudiants");
      await expect(page.locator("#r-mode")).toHaveText("Un message par destinataire");
      await expect(page.locator("#r-exp")).toHaveText(adresse);
      await expect(page.locator('#rythme .option[data-preset="standard"]')).toHaveClass(/actif/);
      await page.locator('#rythme .option[data-preset="soutenu"]').click();
      await expect(page.locator("#r-mode")).toHaveText("Messages groupés de 25 (copie cachée)");
      await expect(page.locator("#r-rythme")).toHaveText("1 000 destinataires par jour");
      await page.locator('#rythme .option[data-preset="prudent"]').click();
      await expect(page.locator("#r-rythme")).toHaveText("150 destinataires par jour");
      await capture(page, f.nom, "6-envoi");

      await page.locator("#lancer-envoi").click();
      await expect(page.locator("#e-envoye")).toBeVisible();
      await expect(page.locator(".demo-seulement")).toBeVisible();
      await expect.poll(() => nombre(page, "#envoyes"), { timeout: 5000 }).toBeGreaterThan(0);
      await capture(page, f.nom, "7-envoye");

      // Rechargé : toujours connecté, et la recherche terminée est proposée.
      await page.reload();
      await expect(page.locator("#e-objectif")).toBeVisible();
      await expect(page.locator("#reprise")).toBeVisible();
      await expect(page.locator("#reprise-texte")).toContainText("100 étudiants identifiés");
      expect(erreurs, erreurs.join("\n")).toEqual([]);
    });

    test(`paramètres (${f.nom})`, async ({ page }) => {
      await connecter(page, email("param"));
      await page.locator("#ouvrir-parametres").click();
      await expect(page.locator("#e-parametres")).toBeVisible();
      await expect(page.locator('#presets .option[data-preset="standard"]')).toHaveClass(/actif/);
      await expect(page.locator("#p-taille")).toHaveValue("1");
      await expect(page.locator("#capacite")).toContainText("400 destinataires par jour");
      await page.locator("#p-taille").fill("60");
      await page.locator("#p-heure").fill("10");
      await page.locator("#p-jour").fill("1500");
      await page.locator("#p-debut").fill("9");
      await page.locator("#p-fin").fill("17");
      await expect(page.locator("#alerte-param")).toBeVisible();
      await expect(page.locator("#alerte-param")).toContainText("50 destinataires par message");
      await expect(page.locator("#alerte-param")).toContainText("1 000 destinataires par jour");
      await expect(page.locator("#presets .option.actif")).toHaveCount(0); // personnalisé
      await expect(page.locator("#capacite")).toContainText("1 500 destinataires par jour. Messages groupés de 60");
      await capture(page, f.nom, "8-parametres");
      await page.locator("#enregistrer-parametres").click();
      await expect(page.locator("#e-objectif")).toBeVisible();

      // Enregistré dans le compte : retrouvé après rechargement.
      await page.reload();
      await page.locator("#ouvrir-parametres").click();
      await expect(page.locator("#p-taille")).toHaveValue("60");
      await expect(page.locator("#p-fin")).toHaveValue("17");
      await page.locator('#presets .option[data-preset="prudent"]').click();
      await expect(page.locator("#p-jour")).toHaveValue("150");
      await expect(page.locator("#alerte-param")).toBeHidden();
      await page.locator("#retour-parametres").click(); // sans enregistrer
      await page.locator("#ouvrir-parametres").click();
      await expect(page.locator("#p-taille")).toHaveValue("60");
    });
  });
}

test("installable : manifeste, icônes et service worker", async ({ page, request }) => {
  const m = await (await request.get(`${URL_APP}/manifest.webmanifest`)).json();
  expect(m.display).toBe("standalone");
  expect(m.icons.map((i: any) => i.sizes)).toEqual(expect.arrayContaining(["192x192", "512x512"]));
  for (const i of m.icons) expect((await request.get(URL_APP + i.src)).headers()["content-type"]).toContain("image/png");
  expect((await request.get(`${URL_APP}/icones/apple-touch-icon.png`)).ok()).toBe(true);
  await page.goto(URL_APP);
  const actif = await page.evaluate(async () => { const r = await navigator.serviceWorker.ready; return !!r.active; });
  expect(actif).toBe(true);
});

test("suspendre puis reprendre sans doublon", async ({ page }) => {
  await connecter(page, email("pause"));
  await page.locator("#objectif").fill("400");
  await page.locator("#start").click();
  await expect(page.locator("#e-recolte")).toBeVisible({ timeout: 20000 });
  await expect.poll(() => nombre(page, "#trouves"), { timeout: 15000 }).toBeGreaterThan(20);
  await page.locator("#pause").click();
  await expect(page.locator("#statut-texte")).toHaveText("Suspendue");
  await expect(page.locator("#vers-envoi")).toBeVisible();
  await page.waitForTimeout(1500);
  const fige = await page.locator("#trouves").textContent();
  await page.waitForTimeout(1500);
  expect(await page.locator("#trouves").textContent()).toBe(fige);
  await page.locator("#pause").click();
  await expect(page.locator("#trouves")).toHaveText("400", { timeout: 60000 });
  const adresses = await page.locator("#liste li .a").allTextContents();
  expect(new Set(adresses).size).toBe(adresses.length);
});

test("mémoire du compte : on ferme en pleine recherche, on reprend sur un autre appareil", async ({ browser }) => {
  const adresse = email("memoire");
  const telephone = await browser.newContext(iphone as any);
  const p1 = await telephone.newPage();
  await connecter(p1, adresse);
  await p1.locator("#objectif").fill("300");
  await p1.locator("#start").click();
  await expect(p1.locator("#e-recolte")).toBeVisible({ timeout: 20000 });
  await expect.poll(() => nombre(p1, "#trouves"), { timeout: 20000 }).toBeGreaterThan(210); // au moins un paquet enregistré
  await telephone.close(); // fermé brutalement en pleine recherche

  const ordinateur = await browser.newContext({ viewport: { width: 1280, height: 860 } });
  const p2 = await ordinateur.newPage();
  await connecter(p2, adresse);
  await expect(p2.locator("#reprise")).toBeVisible();
  await expect(p2.locator("#reprise-texte")).toContainText("sur un objectif de 300");
  const deja = Number((await p2.locator("#reprise-texte").textContent())?.match(/^([\d\s ]+) étudiants/)?.[1].replace(/\D/g, ""));
  expect(deja).toBeGreaterThan(100);
  await p2.locator("#reprendre").click();
  await expect(p2.locator("#e-lancement")).toBeHidden(); // la logique est déjà connue : pas de nouvelle analyse
  await expect(p2.locator("#e-recolte")).toBeVisible();
  await expect(p2.locator("#logique-resume")).toContainText("prenom.nom@demo-ecole.fr");
  await expect(p2.locator("#trouves")).toHaveText("300", { timeout: 30000 });

  // Aucun doublon dans le compte, rien de perdu.
  const code = await p2.evaluate(() => JSON.parse(localStorage.getItem("smartmail") || "{}").code);
  const r = await p2.request.get(`${URL_APP}/v1/destinataires?taille=2000`, { headers: { Authorization: "Bearer " + code } });
  const liste = (await r.json()).personnes.map((p: any) => p.adresse);
  expect(liste).toHaveLength(300);
  expect(new Set(liste).size).toBe(300);
  await ordinateur.close();
});

test("nouvelle recherche : confirmation, puis la liste repart de zéro", async ({ page }) => {
  await connecter(page, email("nouvelle"));
  await page.locator('#raccourcis .puce[data-v="100"]').click();
  await page.locator("#start").click();
  await expect(page.locator("#trouves")).toHaveText("100", { timeout: 40000 });
  await page.reload();
  await expect(page.locator("#reprise")).toBeVisible();
  await expect(page.locator("#reprendre")).toHaveText("Voir les destinataires");
  await page.locator("#nouvelle").click();
  await expect(page.locator("#reprise")).toBeHidden();
  page.once("dialog", (d) => { expect(d.message()).toContain("remplace les 100 destinataires"); d.dismiss(); });
  await page.locator("#start").click();
  await expect(page.locator("#e-objectif")).toBeVisible(); // annulé
  page.once("dialog", (d) => d.accept());
  await page.locator('#raccourcis .puce[data-v="100"]').click();
  await page.locator("#start").click();
  await expect(page.locator("#e-lancement")).toBeVisible();
  await expect(page.locator("#trouves")).toHaveText("100", { timeout: 40000 });
});

test("déconnexion", async ({ page }) => {
  await connecter(page, email("sortie"));
  await page.locator("#ouvrir-parametres").click();
  await page.locator("#deconnexion").click();
  await expect(page.locator("#e-connexion")).toBeVisible();
  await page.reload();
  await expect(page.locator("#e-connexion")).toBeVisible();
});
