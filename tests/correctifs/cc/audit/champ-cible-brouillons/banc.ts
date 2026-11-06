// Banc de l'audit « champ visé, brouillons de l'utilisateur et vol du focus ».
// Smart Mail (extension/) chargée dans Chromium headless, profil temporaire
// distinct par test. AUCUNE requête réelle : les hôtes Outlook sont servis par
// context.route (fausse page owa-monarch.html), tout le reste est bloqué.
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "..", "..", "extension");
export const PAGE = readFileSync(join(ICI, "owa-monarch.html"), "utf8");
export const URL_OUTLOOK = "https://outlook.cloud.microsoft/mail/";
const FAUX_HOTES = new Set(["outlook.cloud.microsoft", "outlook.office.com", "outlook.office365.com", "outlook.live.com"]);

export type Evenement = { type: string; zone: string; confiance: boolean; t: number; origine?: string; inputType?: string; data?: string | null; key?: string; libelle?: string; texte?: string | null };
export type Redaction = { id: string; proprietaire: string; type: string; visible: boolean; a: string[]; aTexte: string; cc: string[]; ccTexte: string; objet: string | null; corps: string };
export type Etat = {
  vue: string; actif: string | null; lu: number | null; redactions: Redaction[]; recherche: string; teams: string;
  teamsEnvoyes: Array<{ texte: string; confiance: boolean }>;
  cal: { titre: string; participants: string[]; participantsTexte: string; description: string };
  contact: { prenom: string; nom: string; mail: string }; signature: string | null; toast: string | null; actifZone: string;
  envoyes: Array<{ id: string; proprietaire: string; type: string; a: string[]; cc: string[]; objet: string; corps: string; clicConfiance: boolean }>;
};

export const test = base.extend<{ context: BrowserContext; worker: Worker; page: Page; requetesBloquees: string[] }>({
  requetesBloquees: async ({}, use) => { await use([]); },
  context: async ({ requetesBloquees }, use) => {
    const dir = mkdtempSync(join(tmpdir(), "audit-champ-cible-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: 1600, height: 1000 },
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    await context.route("**/*", (route) => {
      const u = new URL(route.request().url());
      if (FAUX_HOTES.has(u.hostname) && (u.protocol === "https:")) {
        // ?vitesse=reelle : attentes réelles du Publipostage (sans data-pp-test="rapide").
        const body = u.searchParams.get("vitesse") === "reelle" ? PAGE.replace(' data-pp-test="rapide"', "") : PAGE;
        return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body });
      }
      requetesBloquees.push(u.href);
      return route.abort();
    });
    await use(context);
    await context.close();
    try { rmSync(dir, { recursive: true, force: true }); } catch {}
  },
  worker: async ({ context }, use) => {
    let [w] = context.serviceWorkers();
    if (!w) w = await context.waitForEvent("serviceworker");
    await use(w);
  },
  page: async ({ context, worker }, use) => {
    void worker;
    const page = context.pages()[0] ?? (await context.newPage());
    await use(page);
  },
});
export { expect };

export const ao = (page: Page) => page.locator('[data-testid="ao-panel"]');
export const pp = (page: Page) => page.locator("[data-pp-panneau]");

export async function ouvrir(page: Page, worker: Worker, onglet: "adresses" | "envoi" = "adresses", url = URL_OUTLOOK) {
  await worker.evaluate((o) => chrome.storage.local.set({ "mega-onglet": o }), onglet);
  await page.goto(url);
  await page.bringToFront();
  await page.locator('html[data-faux-pret="1"]').waitFor({ timeout: 20000 });
  await page.locator('html[data-mega-pret="1"]').waitFor({ timeout: 30000 });
  await expect(onglet === "adresses" ? ao(page) : pp(page)).toBeVisible();
}

export const etat = (page: Page) => page.evaluate(() => (window as any).ow.etat()) as Promise<Etat>;
export const journal = (page: Page) => page.evaluate(() => (window as any).__journal.slice()) as Promise<Evenement[]>;
export const viderJournal = (page: Page) => page.evaluate(() => { (window as any).__journal.length = 0; (window as any).__validations.length = 0; });
// Destinataires validés (pastilles créées) depuis le dernier viderJournal, même s'ils ont été effacés ensuite.
export const validations = (page: Page) => page.evaluate(() => (window as any).__validations.slice()) as Promise<Array<{ zone: string; adresse: string; t: number }>>;

// Saisies (insertion ou effacement) faites par l'EXTENSION (événements non fiables), par zone.
// « input » attribués à l'extension (geste synthétique dans la même tâche) ou d'origine inconnue
// (execCommand sans touche : jamais l'utilisateur, dont les frappes Playwright sont des touches réelles).
export async function saisiesExtension(page: Page) {
  return (await journal(page)).filter((e) => e.type === "input" && e.origine !== "utilisateur" && e.zone !== "panneau");
}
export async function saisiesUtilisateur(page: Page) {
  return (await journal(page)).filter((e) => e.type === "input" && e.origine === "utilisateur");
}
export async function zonesTouchees(page: Page) {
  return [...new Set((await saisiesExtension(page)).map((e) => e.zone))];
}
// Clics de l'extension sur la page (hors panneaux).
export async function clicsExtension(page: Page) {
  return (await journal(page)).filter((e) => e.type === "click" && !e.confiance);
}

// Relevé : « Démarrer » (vrai clic : le panneau empêche le vol du focus au mousedown).
export async function demarrerReleve(page: Page) {
  await ao(page).locator("#ao-run").click();
}
export const releveEnCours = (page: Page) => ao(page).locator("#ao-compact-view").isVisible();
// Attend que l'extension ait tapé quelque chose OU que le panneau ait rendu la main avec un message.
export async function attendreDebutOuRefus(page: Page, ms = 20000) {
  const fin = Date.now() + ms;
  while (Date.now() < fin) {
    if ((await saisiesExtension(page)).some((e) => e.inputType === "insertText")) return "saisie";
    const n = ((await ao(page).locator("#ao-notice").textContent()) || "").trim();
    if (n && !(await releveEnCours(page))) return "refus";
    await page.waitForTimeout(100);
  }
  return "rien";
}
// Pause demandée (crochet de test « ao-stop », équivalent du bouton Pause), puis fin du passage.
export async function arreterReleve(page: Page) {
  await page.evaluate(() => document.dispatchEvent(new Event("ao-stop")));
  await expect(ao(page).locator("#ao-compact-view")).toBeHidden({ timeout: 20000 });
}
export async function attendreRecherches(page: Page, n: number, ms = 60000) {
  // n recherches tapées (insertions de l'extension), quelle que soit la zone.
  await expect.poll(async () => (await saisiesExtension(page)).filter((e) => e.inputType === "insertText").length, { timeout: ms }).toBeGreaterThanOrEqual(n);
}
export const noticeReleve = async (page: Page) => ((await ao(page).locator("#ao-notice").textContent()) || "").trim();

// Envoi : remplit le panneau et lance.
export async function lancerEnvoi(page: Page, o: { dest: string; objet?: string; texte?: string; pmin?: number; pmax?: number }) {
  const p = pp(page);
  await p.locator("#pp-dest").fill(o.dest);
  await p.locator("#pp-objet").fill(o.objet ?? "Invitation");
  await p.locator("#pp-texte").fill(o.texte ?? "Bonjour,\n\nVoici l'invitation.\nÀ bientôt");
  await p.locator("#pp-pmin").fill(String(o.pmin ?? 1));
  await p.locator("#pp-pmax").fill(String(o.pmax ?? 1));
  await p.locator("#pp-go").click();
}
export async function attendreFinEnvoi(page: Page, ms = 60000) {
  // L'envoi a démarré (vue « en cours ») ou a été refusé d'emblée (avis), puis il rend la main avec un avis.
  await expect.poll(async () => {
    const live = await pp(page).locator("#pp-live").isVisible();
    const n = ((await pp(page).locator("#pp-notice").textContent()) || "").trim();
    return !live && n !== "";
  }, { timeout: ms }).toBe(true);
  return ((await pp(page).locator("#pp-notice").textContent()) || "").trim();
}
// Relevé lancé depuis la boîte de réception (l'extension ouvre SA rédaction), n recherches faites.
export async function releveLance(page: Page, worker: Worker, n = 2) {
  await ouvrir(page, worker);
  await page.evaluate(() => (window as any).ow.afficherMessage(null));
  await demarrerReleve(page);
  await attendreRecherches(page, n);
  const e = await etat(page);
  const r = e.redactions.find((x) => x.proprietaire === "extension");
  if (!r) throw new Error("le relevé n'a pas ouvert sa propre rédaction");
  return r.id;
}
// Réglages enregistrés du relevé (clé « reglages » du stockage de l'extension), avant chargement de la page.
export async function reglagesReleve(worker: Worker, r: Record<string, unknown>) {
  await worker.evaluate((x) => chrome.storage.local.set({ reglages: JSON.stringify(x) }), r);
}
// La zone qui a le focus devient-elle celle visée (motif) dans le délai ?
export async function attendreFocus(page: Page, motif: RegExp, ms: number) {
  const fin = Date.now() + ms;
  while (Date.now() < fin) {
    if (motif.test((await etat(page)).actifZone)) return true;
    await page.waitForTimeout(80);
  }
  return false;
}
export const contenuUtilisateur = (e: Etat) => ({
  redactions: e.redactions.filter((r) => r.proprietaire === "utilisateur").map((r) => ({ id: r.id, a: r.a, aTexte: r.aTexte, cc: r.cc, ccTexte: r.ccTexte, objet: r.objet, corps: r.corps })),
  recherche: e.recherche, teams: e.teams, cal: e.cal, contact: e.contact, signature: e.signature, toast: e.toast,
});
export const envoyes = async (page: Page) => (await etat(page)).envoyes;
