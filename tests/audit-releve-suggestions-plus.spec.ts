// Audit « Relevé : variantes de la liste de suggestions » — compléments :
// latences longues, pause pendant une réponse lente, Gmail plein écran par-dessus
// un e-mail affiché, rédaction Outlook en fenêtre (dialogue) avec portail.
import { test, expect, ao, url, ouvrir, demarrer, attendreDebut, attendreFin, etat, releve, interdites, pauseApres } from "./audit/releve-suggestions/banc";
import { PETIT, adresses } from "./audit/releve-suggestions/carnet.js";

test.describe.configure({ mode: "parallel" });
test.setTimeout(420000);
const TOUS = adresses(PETIT);

test("X01 suggestions au bout de 10 s, réponses périmées affichées : rien de faux, rien validé, rien laissé", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, url("outlook.office.com", "/mail/inbox", { cap: 5, lot: 5, delai: 10000, perime: true, portail: true }));
  await page.locator("#champ-a").click({ position: { x: 8, y: 8 } });
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await page.waitForTimeout(11000); // les dernières réponses arrivent après la fin
  const e = await etat(page), r = await releve(worker);
  expect(e.violations).toEqual([]);
  expect(interdites(e.keys)).toEqual([]);
  expect(e.champA).toBe("");
  expect(e.pastilles).toBe(0);
  expect(r.txt.filter((m) => !TOUS.includes(m))).toEqual([]);
});

test("X02 pause pendant une réponse lente (3 s) : champ vidé tout de suite, la réponse tardive ne déclenche rien", async ({ banc }) => {
  const { page } = banc;
  await ouvrir(page, url("outlook.office.com", "/mail/inbox", { cap: 5, lot: 5, delai: 3000, perime: true, portail: true }));
  await pauseApres(page, 3);
  await page.locator("#champ-a").click({ position: { x: 8, y: 8 } });
  await demarrer(page);
  await attendreDebut(page);
  const msg = await attendreFin(page, 120000);
  expect(msg).toContain("En pause");
  expect((await etat(page)).champA).toBe("");
  await page.waitForTimeout(4000);
  const e = await etat(page);
  expect(e.violations).toEqual([]);
  expect(e.champA).toBe("");
  expect(e.pastilles).toBe(0);
  expect(interdites(e.keys)).toEqual([]);
});

test("X03 (plausible) Gmail plein écran par-dessus une lettre d'information (lien « Voir plus ») : jamais cliqué", async ({ banc }) => {
  const { page, worker } = banc;
  await page.setViewportSize({ width: 1440, height: 900 });
  await ouvrir(page, url("mail.google.com", "/mail/u/0/", { redaction: "plein-ecran", vue: "message", liens: ["Voir plus"] }));
  await page.locator("#champ-to").click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  const e = await etat(page);
  // BUG : même cause que G08 (clickMore fouille tout le document) : le lien de
  // l'e-mail, sous le voile de la rédaction plein écran, est cliqué.
  expect(e.violations.filter((v) => v.startsWith("clic-lien-email")), "aucun lien d'e-mail cliqué").toEqual([]);
  expect((await releve(worker, "gmail")).txt).toEqual(TOUS);
});

test("X04 Outlook : rédaction en fenêtre ([role=dialog]) avec liste en portail hors du dialogue", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, url("outlook.office.com", "/mail/inbox", { initial: 0, lot: 3, portail: true, redaction: "dialogue", ariaOption: true }));
  await page.locator("#champ-a").click({ position: { x: 8, y: 8 } });
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  const e = await etat(page);
  expect(e.violations).toEqual([]);
  expect(e.champA).toBe("");
  expect((await releve(worker)).txt).toEqual(TOUS);
});
