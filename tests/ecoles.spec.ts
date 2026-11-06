// L'application de bout en bout sur 8 écoles fictives aux conventions différentes
// (serveur/app/ecoles.py). Règle d'or : aucun non-étudiant dans le compte.
//   APP_URL=http://127.0.0.1:8766 node node_modules/@playwright/test/cli.js test tests/ecoles.spec.ts
import { test, expect, type Page } from "@playwright/test";

const URL_APP = process.env.APP_URL ?? "http://127.0.0.1:8766";
const ECOLES = ["prenom-nom", "numeros", "sous-domaine", "intitules-seuls", "nom-prenom", "multi-domaines", "anglophone", "pieges"];
const OBJECTIF = 150;

async function verite(page: Page, ecole: string) {
  const v = new Map<string, string>();
  for (let depart: number | null = 0; depart != null; ) {
    const r = await (await page.request.get(`${URL_APP}/demo/annuaire?ecole=${ecole}&taille=500&depart=${depart}`)).json();
    for (const p of r.personnes) v.set(p.adresse.toLowerCase(), p.verite);
    depart = r.suivant;
  }
  return v;
}

for (const ecole of ECOLES) {
  test(`école « ${ecole} » : objectif atteint, aucun intrus`, async ({ page }) => {
    const erreurs: string[] = [];
    page.on("pageerror", (e) => erreurs.push(String(e)));
    await page.goto(`${URL_APP}/?ecole=${ecole}`);
    await page.locator("#email-demo").fill(`${ecole}-${Date.now()}@demo-ecole.fr`);
    await page.locator('#form-demo button[type="submit"]').click();
    await expect(page.locator("#e-objectif")).toBeVisible();
    await page.locator("#objectif").fill(String(OBJECTIF));
    await page.locator("#start").click();
    await expect(page.locator("#e-recolte")).toBeVisible({ timeout: 30000 });
    await expect(page.locator("#statut-texte")).toHaveText(/Objectif atteint|Annuaire entièrement parcouru/, { timeout: 90000 });

    const code = await page.evaluate(() => JSON.parse(localStorage.getItem("smartmail") || "{}").code);
    const liste = (await (await page.request.get(`${URL_APP}/v1/destinataires?taille=2000`, { headers: { Authorization: "Bearer " + code } })).json()).personnes;
    const v = await verite(page, ecole);
    const intrus = liste.filter((p: any) => v.get(p.adresse) !== "etudiant").map((p: any) => `${p.adresse} (${v.get(p.adresse)}, ${p.source})`);
    expect(intrus, intrus.join("\n")).toEqual([]);
    expect(liste.length).toBe(OBJECTIF);
    expect(new Set(liste.map((p: any) => p.adresse)).size).toBe(liste.length);
    expect(await page.locator("#trouves").textContent()).toBe(String(OBJECTIF));
    expect(erreurs).toEqual([]);
  });
}
