// Audit « Relevé : variantes de la liste de suggestions » — partie 1 : la suite
// d'origine de l'extension Adresses (hellohello-main/extension/tests/e2e/
// extension.spec.mjs) PORTÉE contre Smart Mail, avec ses fausses pages d'origine
// (fakemail/outlook.html, gmail.html, owa-reel.html, book.js, lues sans copie).
// Différences d'adaptation (Smart Mail) :
//  - pages servies par context.route sous https://outlook.office.com/orig/… et
//    https://mail.google.com/orig/… (plus de serveur localhost:5178) ;
//  - plus de test-config.js : l'extension tourne à son rythme RÉEL (≈1,6 s par
//    recherche), d'où des délais plus longs ;
//  - plus de démarrage automatique : l'utilisateur clique dans le champ puis sur
//    « Démarrer » (ex « Démarrer ici ») ; « Reprendre » remplace « Reprendre ici » ;
//  - la croix ne retire plus le panneau : elle le MASQUE (mega.js).
import { test, expect, ao, ouvrir, demarrer, attendreDebut, attendreFin, releve, stockage, interdites, pauseApres, simuler, Planif, Moteur } from "./audit/releve-suggestions/banc";
import { ask, BOOK } from "../../Adresses Outlook/hellohello-main/extension/tests/e2e/fakemail/book.js";

test.describe.configure({ mode: "parallel" });
test.setTimeout(480000);

const SIM = simuler((q: string) => ask(q));
const OUTLOOK = "https://outlook.office.com/orig/outlook.html";
const GMAIL = "https://mail.google.com/orig/gmail.html";
const OWA_REEL = "https://outlook.office.com/orig/owa-reel.html";
const JKEY = "reprise-outlook";
const typedDe = (page: any) => page.evaluate(() => (window as any).__typed as string[]);
const keysDe = (page: any) => page.evaluate(() => ((window as any).__keys || []) as string[]);

// Journal « façon application ps1 » arrêté après n recherches (vrais Planif et
// Moteur.journalLine de Smart Mail), comme buildAppJournal d'origine.
function journalApp(n: number) {
  const plan = new Planif("");
  const head = { Started: "2026-10-14T09:00:00", Base: "C:\\Users\\user\\Desktop\\adresses-outlook-2026-10-14-0900", Seconds: 42, Source: "outlook", GmailAccount: "", Field: "to", Platform: "app", Prefix: "", Keep: "", Exclude: "", Dedupe: true, Format: "txt", Digits: true };
  const lignes = [JSON.stringify(head)];
  for (let i = 0; i < n; i++) {
    const q = plan.next(); if (q === null) break;
    const list = ask(q), keep = list.map(() => true);
    plan.report(q, list.map((x: any) => x.name), list.map((x: any) => x.mail), keep);
    lignes.push(Moteur.journalLine(q, list, keep));
  }
  return lignes.join("\n") + "\n";
}

test("P1 Outlook : l'utilisateur tape une lettre dans « À » puis « Démarrer » → séquence = planificateur, journal effacé, export", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await page.locator('input[aria-label="To"]').click();
  await page.locator('input[aria-label="To"]').pressSequentially("a", { delay: 5 }); // l'ancien geste d'armement
  await page.waitForTimeout(1500);
  await expect(ao(page).locator("#ao-compact-view"), "Smart Mail : plus de démarrage automatique").toBeHidden();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/);
  expect(await typedDe(page)).toEqual(SIM.seq);
  expect(interdites(await keysDe(page)), "touche de validation tapée").toEqual([]);
  const r = await releve(worker);
  expect(r.journal, "journal effacé à la fin").toBeUndefined();
  expect(r.txt).toEqual(SIM.distinct);
  expect(await page.locator('input[aria-label="To"]').inputValue(), "rien laissé dans « À »").toBe("");
  const [dl] = await Promise.all([page.waitForEvent("download"), ao(page).locator("#ao-export").click()]);
  expect(dl.suggestedFilename()).toMatch(/^adresses-outlook-\d{4}-\d{2}-\d{2}\.txt$/);
});

test("P2 Gmail : l'utilisateur tape une lettre dans « À » puis « Démarrer » → séquence = planificateur", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, GMAIL);
  await page.locator('div[name="to"] input').click();
  await page.locator('div[name="to"] input').pressSequentially("a", { delay: 5 });
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/);
  expect(await typedDe(page)).toEqual(SIM.seq);
  expect(interdites(await keysDe(page))).toEqual([]);
  expect((await releve(worker, "gmail")).txt).toEqual(SIM.distinct);
  expect(await page.locator('div[name="to"] input').inputValue()).toBe("");
});

test("P3 Gmail : curseur dans « À » sans rien taper, « Démarrer » → séquence = planificateur", async ({ banc }) => {
  const { page } = banc;
  await ouvrir(page, GMAIL);
  await page.locator('div[name="to"] input').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/);
  expect(await typedDe(page)).toEqual(SIM.seq);
  expect(interdites(await keysDe(page))).toEqual([]);
});

test("P4 taper dans la barre de recherche n'arme pas et ne démarre pas ; « Démarrer » depuis la recherche vise « À », jamais la recherche", async ({ banc }) => {
  const { page } = banc;
  await ouvrir(page, OUTLOOK);
  await page.locator('input[aria-label="Rechercher"]').click();
  await page.locator('input[aria-label="Rechercher"]').pressSequentially("ab", { delay: 5 });
  await page.waitForTimeout(1500);
  await expect(page.locator('[data-testid="ao-panel"][data-ao-armed="1"]'), "la recherche ne doit pas armer").toHaveCount(0);
  expect(await typedDe(page), "aucun passage ne démarre").toEqual([]);
  // Smart Mail : « Démarrer » alors que le curseur est dans la recherche.
  await pauseApres(page, 3);
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  expect(await page.locator('input[aria-label="Rechercher"]').inputValue(), "la recherche n'est jamais touchée").toBe("ab");
  expect((await typedDe(page)).length).toBeGreaterThanOrEqual(3); // c'est « À » qui a reçu les recherches
  expect(await page.locator('input[aria-label="To"]').inputValue()).toBe("");
});

test("P5 curseur dans Cc + « Démarrer » → récolte sur Cc, « À » jamais touché", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await page.evaluate(() => { (window as any).__toInputs = 0; document.querySelector('input[aria-label="To"]')!.addEventListener("input", () => (window as any).__toInputs++); });
  await page.locator('input[aria-label="Cc"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/);
  expect(await typedDe(page)).toEqual(SIM.seq);
  expect(await page.evaluate(() => (window as any).__toInputs), "rien tapé dans « À »").toBe(0);
  expect((await releve(worker)).txt).toEqual(SIM.distinct);
  expect(await page.locator('input[aria-label="Cc"]').inputValue()).toBe("");
});

test("P6 (différence) curseur dans « À » sans rien taper : AUCUN démarrage automatique dans Smart Mail", async ({ banc }) => {
  const { page } = banc;
  await ouvrir(page, OUTLOOK);
  await page.locator('input[aria-label="To"]').click();
  await page.waitForTimeout(5000); // l'original démarrait seul après ~2 s (400 ms en test)
  await expect(ao(page).locator("#ao-compact-view")).toBeHidden();
  expect(await typedDe(page)).toEqual([]);
  expect(await page.locator('input[aria-label="To"]').inputValue()).toBe("");
});

test("P7 OWA réel : clique « Afficher plus de résultats » (pied de page, liste vide au départ), boîte de réception jamais lue, leurre jamais visé", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OWA_REEL);
  await page.locator("#to1").click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/);
  expect(await typedDe(page)).toEqual(SIM.seq);
  expect(await page.evaluate(() => (window as any).__moreClicks)).toBeGreaterThan(0);
  const r = await releve(worker);
  expect(r.txt.join("\n")).not.toContain("noreply@email.apple.com");
  expect(r.txt).toEqual(SIM.distinct); // lot COMPLET capté, depuis une liste vide
  expect(await page.evaluate(() => (window as any).__leurreFocus), "le champ leurre n'a jamais été visé").toBe(0);
  expect(await page.locator("#to1").textContent()).toBe("");
});

test("P8 arrêt immédiat si un destinataire est validé par erreur (jamais d'envoi) — et rien laissé dans « À »", async ({ banc }) => {
  const { page } = banc;
  await ouvrir(page, OUTLOOK);
  await page.evaluate(() => {
    const tick = () => { if (((window as any).__typed || []).length >= 3) document.dispatchEvent(new Event("ao-make-chip")); else requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  });
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  const msg = await attendreFin(page, 60000);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/warn/);
  expect(msg).toContain("Arrêté");
  expect(interdites(await keysDe(page))).toEqual([]);
  expect((await typedDe(page)).length).toBeLessThan(6); // arrêté aussitôt
  // BUG : à l'arrêt « destinataire validé », le moteur ne vide pas « À » : la
  // dernière recherche tapée (ex. « d ») reste dans le champ (R2). Cause :
  // moteur.js n'efface en fin de passage que si focus() réussit, or focus()
  // échoue justement dès qu'une pastille est apparue.
  expect.soft(await page.locator('input[aria-label="To"]').inputValue(), "rien laissé dans « À » après l'arrêt").toBe("");
  // BUG : le panneau dit « le curseur a quitté le champ « À » » : il ne prévient
  // pas qu'un destinataire a été validé dans le brouillon (à retirer avant tout
  // envoi). Cause : moteur.js confond l'échec de focus() « pastille apparue » avec
  // une perte de focus (R6 : s'arrêter ET dire la vraie raison).
  expect.soft(msg, "le message signale le destinataire validé par erreur").toMatch(/destinataire/i);
});

test("P9 arrêt puis reprise après rechargement : finit sans retaper, journal effacé à la fin", async ({ banc }) => {
  test.setTimeout(600000);
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await pauseApres(page, 5);
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page, 60000);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/warn/);
  expect((await releve(worker)).journal, "journal gardé après la pause").toBeTruthy();
  const typed1 = await typedDe(page);
  expect(typed1.length).toBeGreaterThan(0);
  expect(typed1).toEqual(SIM.seq.slice(0, typed1.length));
  expect(await page.locator('input[aria-label="To"]').inputValue(), "pause : champ vidé").toBe("");

  await page.reload();
  await page.bringToFront();
  await ao(page).and(page.locator('[data-ao-ready="1"]')).waitFor({ state: "attached" });
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached" });
  await expect(ao(page).locator("#ao-resume")).toBeVisible();
  await expect(ao(page).locator("#ao-run")).toHaveText("Reprendre");
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/);
  const typed2 = await typedDe(page);
  const total = typed1.length + typed2.length;
  expect(total).toBeGreaterThanOrEqual(SIM.seq.length);
  expect(total).toBeLessThanOrEqual(SIM.seq.length + 1);
  const r = await releve(worker);
  expect(r.journal, "journal effacé à la fin de la reprise").toBeUndefined();
  expect(r.txt).toEqual(SIM.distinct);
});

test("P10 import : une liste d'adresses rejoint l'historique (compteur + Export)", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await ao(page).locator("#ao-file").setInputFiles({
    name: "export.txt", mimeType: "text/plain",
    buffer: Buffer.from("MARTIN Alice <alice.martin@essca.fr>\n- DURAND Bob <bob.durand@essca.fr>\ncarla@essca.fr\nALICE.MARTIN@ESSCA.FR\n"),
  });
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/ok/, { timeout: 15000 });
  await expect(ao(page).locator("#ao-export")).toBeEnabled();
  await expect(ao(page).locator("#ao-distinct")).toHaveText(/3 adresses distinctes/);
  const all = await stockage(worker);
  expect(String(all["base-outlook"] || "")).toContain("alice.martin@essca.fr");
  expect(String(all["base-outlook"] || "")).toContain("bob.durand@essca.fr");
  expect(String(all["base-outlook"] || "")).toContain("carla@essca.fr");
});

test("P11 handoff : importer un point de reprise « façon ps1 » (en-tête Field/Platform, Base Windows)", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  const US = "\x1f";
  const head = JSON.stringify({ Started: "2026-10-14T10:00:00", Base: "C:\\Users\\user\\Desktop\\adresses-outlook-2026-10-14-1000", Seconds: 12, Source: "outlook", GmailAccount: "", Field: "to", Platform: "app", Prefix: "", Keep: "", Exclude: "", Dedupe: true, Format: "txt", Digits: true });
  const journal = [head, "a\tAlice Martin" + US + "alice.martin@essca.fr" + US + "1", "b\tBob Durand" + US + "bob.durand@essca.fr" + US + "1"].join("\n") + "\n";
  await ao(page).locator("#ao-file").setInputFiles({ name: "reprise-outlook.jsonl", mimeType: "text/plain", buffer: Buffer.from(journal) });
  await expect(ao(page).locator("#ao-resume")).toBeVisible({ timeout: 15000 });
  const all = await stockage(worker);
  expect(all["reprise-outlook"]).toBeTruthy();
  expect(String(all["base-outlook"] || "")).toContain("alice.martin@essca.fr");
});

test("P12 handoff app→extension : reprend sans rescanner (seules les recherches restantes sont tapées)", async ({ banc }) => {
  const { page, worker } = banc;
  const N = 12;
  await ouvrir(page, OUTLOOK);
  await ao(page).locator("#ao-file").setInputFiles({ name: "reprise-outlook.txt", mimeType: "text/plain", buffer: Buffer.from(journalApp(N)) });
  await expect(ao(page).locator("#ao-resume")).toBeVisible({ timeout: 15000 });
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveText(/Terminé/);
  expect(await typedDe(page)).toEqual(SIM.seq.slice(N));
  expect((await releve(worker)).txt).toEqual(SIM.distinct);
});

test("P13 (différence) croix : masque le panneau (au lieu de le retirer), rien ne s'arme ni ne démarre, revient au rechargement", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await page.evaluate(() => { (window as any).__began = 0; document.addEventListener("ao-run-begin", () => (window as any).__began++); });
  await ao(page).locator("#ao-close").click();
  await expect(ao(page)).toBeHidden();
  await expect(page.locator('[data-testid="ao-panel"]'), "Smart Mail : le panneau reste dans la page, masqué").toHaveCount(1);
  await page.locator('input[aria-label="To"]').click();
  await page.locator('input[aria-label="To"]').pressSequentially("a", { delay: 5 });
  await page.waitForTimeout(3000);
  expect(await page.evaluate(() => (window as any).__began), "aucun passage après fermeture").toBe(0);
  expect((await stockage(worker))["actif-outlook"]).toBeUndefined();
  await page.reload();
  await ao(page).and(page.locator('[data-ao-ready="1"]')).waitFor({ state: "attached" });
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached" });
  await expect(ao(page).locator("#ao-close")).toBeVisible();
});

test("P14 (différence) croix PENDANT un passage : Smart Mail masque le panneau et le relevé CONTINUE jusqu'au bout", async ({ banc }) => {
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await page.evaluate(() => {
    const tick = () => { if (((window as any).__typed || []).length >= 5) (document.querySelector('[data-testid="ao-panel"]') as any).shadowRoot.getElementById("ao-close").click(); else requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  });
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await expect(ao(page)).toBeHidden({ timeout: 30000 });
  // Original : arrêt propre après la recherche en cours, journal gardé.
  // Smart Mail : le panneau est seulement masqué ; le relevé va au bout.
  await expect.poll(async () => (await stockage(worker))["actif-outlook"], { timeout: 300000, intervals: [2000] }).toBeUndefined();
  expect(await typedDe(page)).toEqual(SIM.seq);
  expect((await releve(worker)).journal).toBeUndefined();
  expect(await page.locator('input[aria-label="To"]').inputValue()).toBe("");
});

test("P15 reprise portable : exporter le point de reprise, l'oublier, le réimporter, continuer", async ({ banc }) => {
  test.setTimeout(600000);
  const { page, worker } = banc;
  await ouvrir(page, OUTLOOK);
  await pauseApres(page, 5);
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page, 60000);
  await expect(ao(page).locator("#ao-notice")).toHaveClass(/warn/);
  await ao(page).locator("#ao-gear").click();
  const [dl] = await Promise.all([page.waitForEvent("download"), ao(page).locator("#ao-save-resume").click()]);
  const fichier = await dl.path();
  await ao(page).locator("#ao-gear").click();
  await ao(page).locator("#ao-fresh").click();
  await expect(ao(page).locator("#ao-resume")).toBeHidden();
  await page.reload();
  await ao(page).and(page.locator('[data-ao-ready="1"]')).waitFor({ state: "attached" });
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached" });
  await expect(ao(page).locator("#ao-resume")).toBeHidden();
  await ao(page).locator("#ao-file").setInputFiles(fichier);
  await expect(ao(page).locator("#ao-resume")).toBeVisible({ timeout: 15000 });
  await page.locator('input[aria-label="To"]').click();
  await demarrer(page);
  await attendreDebut(page);
  await attendreFin(page);
  await expect(ao(page).locator("#ao-notice")).toHaveText(/Terminé/);
  const r = await releve(worker);
  expect(r.journal).toBeUndefined();
  expect(r.txt).toEqual(SIM.distinct);
});

test("P16 sans curseur dans un champ (Gmail) : « Démarrer » refuse proprement, rien n'est tapé", async ({ banc }) => {
  const { page } = banc;
  await ouvrir(page, GMAIL);
  await demarrer(page);
  await expect(ao(page).locator("#ao-notice")).toContainText("Clique dans le champ", { timeout: 15000 });
  await expect(ao(page).locator("#ao-compact-view")).toBeHidden();
  expect(await typedDe(page)).toEqual([]);
});
