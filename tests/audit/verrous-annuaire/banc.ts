// Banc de l'audit « verrous entre relevé et envoi, annuaire commun et listes
// « Déjà envoyés » sous contrainte ». Smart Mail (extension/) chargée dans
// Chromium headless, un profil temporaire DISTINCT par test. AUCUNE requête
// réelle : https://outlook.office.com/mail/… est servi par context.route
// (fausse page outlook.html de ce dossier), POST /__test/envoi est intercepté
// pour tenir le journal GLOBAL des mails « envoyés » (tous onglets, rechargements
// compris), tout le reste est bloqué (route.abort).
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker, type Locator } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "extension");
const MODELE = readFileSync(join(ICI, "outlook.html"), "utf8");

export type Envoi = { onglet: string; n: number; a: string[]; cc: string[]; cci: string[]; objet: string; corps: string; html: string; importance: string; t: number };
export type Journal = {
  frappes: Array<{ zone: string; key: string; confiance: boolean; t: number }>;
  saisies: Array<{ zone: string; inputType: string; data: string | null; confiance: boolean; t: number }>;
  clics: Array<{ cible: string; id?: string; confiance: boolean; t: number }>;
  focus: Array<{ zone: string; t: number }>;
  redactions: number;
};
export type Banc = { context: BrowserContext; sw: Worker; envois: Envoi[]; bloquees: string[] };

function pageHtml(u: URL): string {
  const attrs = u.searchParams.get("rapide") === "1" ? ' data-pp-test="rapide"' : "";
  return MODELE.replace('<html lang="fr">', `<html lang="fr"${attrs}>`);
}

export const test = base.extend<{ banc: Banc }>({
  // eslint-disable-next-line no-empty-pattern
  banc: async ({}, use) => {
    const dir = mkdtempSync(join(tmpdir(), "audit-verrous-annuaire-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: 1280, height: 860 },
      acceptDownloads: true,
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    const envois: Envoi[] = [];
    const bloquees: string[] = [];
    // Les routes enregistrées en dernier passent en premier : d'abord « tout bloquer »…
    await context.route("**/*", (route) => { bloquees.push(route.request().url()); return route.abort(); });
    // … puis la fausse page Outlook (seul hôte servi).
    await context.route("https://outlook.office.com/**", (route) => {
      const u = new URL(route.request().url());
      if (u.pathname.startsWith("/__test/envoi")) {
        try { envois.push(JSON.parse(route.request().postData() || "{}")); } catch { /* journal */ }
        return route.fulfill({ status: 204, body: "" });
      }
      if (u.pathname.startsWith("/mail")) return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: pageHtml(u) });
      bloquees.push(u.href);
      return route.abort();
    });
    let [sw] = context.serviceWorkers();
    if (!sw) sw = await context.waitForEvent("serviceworker");
    await use({ context, sw, envois, bloquees });
    await context.close();
    try { rmSync(dir, { recursive: true, force: true }); } catch { /* profil temporaire */ }
  },
});
export { expect };
export type { Page, Locator, BrowserContext, Worker };

// ------------------------------------------------------------ pages
export const ao = (p: Page) => p.locator('[data-testid="ao-panel"]');
export const pp = (p: Page) => p.locator("[data-pp-panneau]");
export const ongletMega = (p: Page, o: "adresses" | "envoi") => p.locator(`[data-testid="ao-panel"]:visible, [data-pp-panneau]:visible`).locator(`.mega-tab[data-onglet="${o}"]`);

export type OptionsOnglet = { onglet?: string; rapide?: boolean; delai?: number; vue?: "adresses" | "envoi"; page?: Page };
export function urlOutlook(o: OptionsOnglet = {}): string {
  const q = new URLSearchParams();
  q.set("onglet", o.onglet ?? "A");
  if (o.rapide !== false) q.set("rapide", "1");
  if (o.delai != null) q.set("delai", String(o.delai));
  return "https://outlook.office.com/mail/inbox/?" + q.toString();
}
// Ouvre un onglet Outlook (fausse page) et attend que Smart Mail soit prêt.
export async function ouvrirOnglet(b: Banc, o: OptionsOnglet = {}): Promise<Page> {
  const page = o.page ?? (b.context.pages().find((x) => x.url() === "about:blank") ?? (await b.context.newPage()));
  await page.goto(urlOutlook(o));
  await page.bringToFront();
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await installerSonde(page);
  if (o.vue) await montrerOnglet(page, o.vue);
  return page;
}
export async function recharger(page: Page): Promise<void> {
  await page.reload();
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await installerSonde(page);
}
export async function montrerOnglet(page: Page, o: "adresses" | "envoi"): Promise<void> {
  const visible = (await pp(page).isVisible()) ? pp(page) : ao(page);
  await visible.locator(`.mega-tab[data-onglet="${o}"]`).click();
  await expect(o === "envoi" ? pp(page) : ao(page)).toBeVisible();
}

// Sonde dans la page (monde principal) : état des deux moteurs lu dans les
// panneaux (shadow DOM ouvert). Des MutationObserver sur « hidden » de la vue
// « Relevé en cours » (ao-compact-view) et de la vue d'envoi (pp-live) notent
// CHAQUE changement : « ensemble » compte les instants où le relevé ET l'envoi
// sont en cours à la fois (verrou mutuel rompu), sans dépendre d'un échantillonnage.
// « arme » : le panneau Adresses a été armé (data-ao-armed) ; debutsReleve : événements ao-run-begin.
export type Sonde = { ensemble: number; releve: number; envoi: number; arme: number; debutsReleve: number[]; debutsEnvoi: number[]; noticesEnvoi: number; noticesReleve: number };
export async function installerSonde(page: Page): Promise<void> {
  await page.evaluate(() => {
    const w = window as unknown as { __sonde?: Sonde };
    if (w.__sonde) return;
    const s = (w.__sonde = { ensemble: 0, releve: 0, envoi: 0, arme: 0, debutsReleve: [] as number[], debutsEnvoi: [] as number[], noticesEnvoi: 0, noticesReleve: 0 });
    document.addEventListener("ao-run-begin", () => s.debutsReleve.push(Date.now()));
    const hA = document.querySelector('[data-testid="ao-panel"]') as HTMLElement, hP = document.querySelector("[data-pp-panneau]") as HTMLElement;
    const a = hA.shadowRoot!.getElementById("ao-compact-view")!, p = hP.shadowRoot!.getElementById("pp-live")!;
    let r0 = !a.hidden, e0 = !p.hidden;
    const maj = () => {
      const r = !a.hidden, e = !p.hidden;
      if (r && !r0) s.releve++;
      if (e && !e0) { s.envoi++; s.debutsEnvoi.push(Date.now()); }
      if (r && e) s.ensemble++;
      r0 = r; e0 = e;
    };
    new MutationObserver(maj).observe(a, { attributes: true, attributeFilter: ["hidden"] });
    new MutationObserver(maj).observe(p, { attributes: true, attributeFilter: ["hidden"] });
    new MutationObserver(() => { if (hA.dataset.aoArmed) s.arme++; }).observe(hA, { attributes: true, attributeFilter: ["data-ao-armed"] });
    new MutationObserver(() => { s.noticesEnvoi++; }).observe(hP.shadowRoot!.getElementById("pp-notice")!, { childList: true, characterData: true, subtree: true, attributes: true });
    new MutationObserver(() => { s.noticesReleve++; }).observe(hA.shadowRoot!.getElementById("ao-notice")!, { childList: true, characterData: true, subtree: true, attributes: true });
    setInterval(() => { if (!a.hidden && !p.hidden) s.ensemble++; }, 10); // filet : échantillon toutes les 10 ms
  });
}
export async function sonde(page: Page): Promise<Sonde> {
  return page.evaluate(() => (window as unknown as { __sonde: Sonde }).__sonde);
}
export async function etat(page: Page): Promise<{ releve: boolean; envoi: boolean; aoNotice: string; ppNotice: string }> {
  return page.evaluate(() => {
    const ra = document.querySelector('[data-testid="ao-panel"]')?.shadowRoot, rp = document.querySelector("[data-pp-panneau]")?.shadowRoot;
    return {
      releve: !!ra && !ra.getElementById("ao-compact-view")!.hidden,
      envoi: !!rp && !rp.getElementById("pp-live")!.hidden,
      aoNotice: ra?.getElementById("ao-notice")?.textContent ?? "",
      ppNotice: rp?.getElementById("pp-notice")?.textContent ?? "",
    };
  });
}
export async function journal(page: Page): Promise<Journal> {
  return page.evaluate(() => (window as unknown as { __journal: Journal }).__journal);
}
export async function cfgPage(page: Page, c: Record<string, unknown>): Promise<void> {
  await page.evaluate((x) => Object.assign((window as unknown as { __cfg: Record<string, unknown> }).__cfg, x), c);
}
export async function envoyesPage(page: Page): Promise<Envoi[]> {
  return JSON.parse((await page.locator("html").getAttribute("data-envoyes")) ?? "[]");
}

// ------------------------------------------------------------ envoi (Publipostage)
export type Envoi0 = { dest?: string[]; objet?: string; texte?: string; pmin?: number; pmax?: number; max?: number; annuaire?: boolean };
export async function preparerEnvoi(page: Page, e: Envoi0): Promise<void> {
  const p = pp(page);
  await expect(p).toBeVisible();
  if (e.annuaire === false) { if (await p.locator("#pp-annuaire-actif").isVisible()) await p.locator("#pp-annuaire-actif").uncheck(); }
  if (e.dest) await p.locator("#pp-dest").fill(e.dest.join("\n"));
  await p.locator("#pp-objet").fill(e.objet ?? "Objet de test");
  await p.locator("#pp-texte").fill(e.texte ?? "Bonjour,\nTexte du message.");
  await p.locator("#pp-pmin").fill(String(e.pmin ?? 0));
  await p.locator("#pp-pmax").fill(String(e.pmax ?? e.pmin ?? 0));
  if (e.max != null) await p.locator("#pp-max").fill(String(e.max));
  await page.waitForTimeout(400); // brouillon et réglages enregistrés (300 ms)
}
const avantLancer = new WeakMap<Page, number>();
// Clique « Lancer l'envoi » / « Reprendre l'envoi » en notant le nombre de notices déjà affichées par le panneau d'envoi.
export async function lancerEnvoi(page: Page): Promise<void> {
  if (!(await pp(page).isVisible())) await montrerOnglet(page, "envoi");
  avantLancer.set(page, (await sonde(page)).noticesEnvoi);
  await pp(page).locator("#pp-go").click();
}
// Fin de l'envoi lancé par lancerEnvoi : une NOUVELLE notice (terminé, pause, limite, refus, arrêt) est affichée et la
// vue d'envoi est refermée. (Une notice restée d'avant, par exemple d'un import, ne compte pas.)
export async function finEnvoi(page: Page, timeout = 120000): Promise<void> {
  const avant = avantLancer.get(page) ?? 0;
  await expect.poll(async () => {
    const [s, e] = await Promise.all([sonde(page), etat(page)]);
    return s.noticesEnvoi > avant && !e.envoi && e.ppNotice.trim() !== "";
  }, { timeout }).toBe(true);
}
export async function pauseEnvoi(page: Page): Promise<void> {
  await pp(page).locator("#pp-pause").click();
  await expect(pp(page).locator("#pp-live")).toBeHidden({ timeout: 60000 });
}
export async function attendreEnvois(b: Banc, n: number, timeout = 60000): Promise<Envoi[]> {
  await expect.poll(() => b.envois.length, { timeout }).toBeGreaterThanOrEqual(n);
  return b.envois;
}
export const destinataires = (b: Banc) => b.envois.map((m) => [...m.a, ...m.cc, ...m.cci].join(",")).flatMap((x) => x.split(",")).filter(Boolean);

// ------------------------------------------------------------ relevé (Adresses)
export async function demarrerReleve(page: Page): Promise<void> {
  if (!(await ao(page).isVisible())) await montrerOnglet(page, "adresses");
  await ao(page).locator("#ao-run").click();
}
export async function attendreReleve(page: Page, timeout = 30000): Promise<void> {
  await expect.poll(async () => (await etat(page)).releve, { timeout }).toBe(true);
}
export async function pauseReleve(page: Page): Promise<void> {
  if (!(await etat(page)).releve) return;
  if (!(await ao(page).isVisible())) await montrerOnglet(page, "adresses");
  await ao(page).locator("#ao-stop").click();
  await expect.poll(async () => (await etat(page)).releve, { timeout: 30000 }).toBe(false);
}

// ------------------------------------------------------------ stockage de l'extension
export async function poser(b: Banc, o: Record<string, unknown>): Promise<void> {
  await b.sw.evaluate((x) => chrome.storage.local.set(x), o);
}
export async function lire<T = unknown>(b: Banc, k: string): Promise<T> {
  return b.sw.evaluate(async (cle) => (await chrome.storage.local.get(cle))[cle], k) as Promise<T>;
}
// Bouton de l'extension (barre de Chrome) : sw.js envoie { mega: "basculer" } à l'onglet cliqué.
// En headless on ne peut pas cliquer l'icône, et sans permission « tabs » l'extension ne voit pas l'URL des
// onglets : on envoie le même message que son gestionnaire à l'onglet Outlook désigné par son rang d'ouverture
// (les id d'onglets croissent avec l'ouverture ; les onglets sans script de contenu ignorent le message).
export async function boutonExtension(b: Banc, page: Page): Promise<void> {
  const rang = b.context.pages().filter((p) => p.url().startsWith("https://outlook.office.com/")).indexOf(page);
  const ok = await b.sw.evaluate(async (r) => {
    const tabs = (await chrome.tabs.query({})).sort((x, y) => (x.id ?? 0) - (y.id ?? 0));
    const vivants: number[] = [];
    for (const t of tabs) {
      try { await chrome.tabs.sendMessage(t.id!, { mega: "ping-audit" }); vivants.push(t.id!); } catch { /* pas de script de contenu */ }
    }
    const id = vivants[r];
    if (id == null) return false;
    await chrome.tabs.sendMessage(id, { mega: "basculer" });
    return true;
  }, rang);
  if (!ok) throw new Error("onglet Outlook introuvable pour le bouton de l'extension");
}
export async function badges(b: Banc): Promise<string[]> {
  return b.sw.evaluate(async () => {
    const out = [await chrome.action.getBadgeText({})];
    for (const t of await chrome.tabs.query({})) out.push(await chrome.action.getBadgeText({ tabId: t.id! }));
    return out.filter(Boolean);
  });
}

// Adresses de test
export const liste = (prefixe: string, n: number, domaine = "exemple.fr") => Array.from({ length: n }, (_, i) => `${prefixe}${i + 1}@${domaine}`);
