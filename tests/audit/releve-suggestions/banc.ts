// Banc commun de l'audit « Relevé : variantes de la liste de suggestions ».
// Smart Mail (extension/, NON modifiée) chargée dans Chromium headless, profil
// temporaire propre à chaque test. TOUTES les pages viennent de context.route :
// outlook.* et mail.google.com sont servis par de fausses pages, tout le reste
// est bloqué (route.abort). Aucune requête réelle, aucun vrai mail.
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync, mkdirSync } from "node:fs";

export const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "extension");
// Suite d'origine de l'extension « Adresses » (fausses pages lues telles quelles).
export const ORIG = "/Users/test/Desktop/Applications/Adresses Outlook/hellohello-main/extension/tests/e2e/fakemail";
const PROFILS = "/private/tmp/claude-502/-Users-test/471010cb-7f57-426b-a93c-e190a5473d31/scratchpad/audit/releve-suggestions/profils";

const lire = (p: string) => readFileSync(p, "utf8");
const HOTES = /^https:\/\/(outlook\.office\.com|outlook\.office365\.com|outlook\.live\.com|outlook\.cloud\.microsoft|mail\.google\.com)\//;

export type Banc = { context: BrowserContext; worker: Worker; page: Page; sorties: string[] };

export const test = base.extend<{ banc: Banc }>({
  // eslint-disable-next-line no-empty-pattern
  banc: async ({}, use) => {
    mkdirSync(PROFILS, { recursive: true });
    const dir = mkdtempSync(join(PROFILS, "p-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: 1600, height: 900 },
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`, "--disable-background-networking", "--disable-component-update", "--no-first-run"],
    });
    context.setDefaultTimeout(20000); // aucun geste ne doit attendre sans fin
    const sorties: string[] = []; // toute requête qui aurait voulu sortir (bloquée)
    await context.route("**/*", (route) => { sorties.push(route.request().url()); return route.abort(); });
    await context.route(HOTES, (route) => {
      const u = new URL(route.request().url());
      const p = u.pathname;
      const js = (body: string) => route.fulfill({ status: 200, contentType: "text/javascript; charset=utf-8", body });
      const html = (body: string) => route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body });
      if (p.endsWith("/favicon.ico")) return route.fulfill({ status: 204, body: "" });
      if (p.endsWith("/carnet.js")) return js(lire(join(ICI, "carnet.js")));
      if (p.endsWith("/book.js")) return js(lire(join(ORIG, "book.js")));
      if (p.startsWith("/orig/")) return html(lire(join(ORIG, p.slice("/orig/".length))));
      if (u.hostname === "mail.google.com") return html(lire(join(ICI, "gmail.html")));
      return html(lire(join(ICI, "owa.html")));
    });
    let [worker] = context.serviceWorkers();
    if (!worker) worker = await context.waitForEvent("serviceworker");
    const page = context.pages()[0] ?? (await context.newPage());
    await use({ context, worker, page, sorties });
    await context.close();
    try { rmSync(dir, { recursive: true, force: true }); } catch {}
  },
});
export { expect };

export const ao = (page: Page) => page.locator('[data-testid="ao-panel"]');

export function url(hote: string, chemin: string, cfg?: object) {
  return `https://${hote}${chemin}` + (cfg ? `?cfg=${encodeURIComponent(JSON.stringify(cfg))}` : "");
}

// Ouvre une fausse page et attend que Smart Mail soit prêt (panneau Adresses ;
// sur Outlook aussi mega.js, qui ne tourne pas sur Gmail).
export async function ouvrir(page: Page, adresse: string) {
  await page.goto(adresse);
  await page.bringToFront();
  await ao(page).and(page.locator('[data-ao-ready="1"]')).waitFor({ state: "attached", timeout: 30000 });
  if (!/mail\.google\.com/.test(adresse)) await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
}

export async function demarrer(page: Page) {
  await ao(page).locator("#ao-run").click();
}
export async function attendreDebut(page: Page, timeout = 20000) {
  await expect(ao(page).locator("#ao-compact-view")).toBeVisible({ timeout });
}
// Fin du passage : vue compacte refermée, message posé.
export async function attendreFin(page: Page, timeout = 300000) {
  await expect(ao(page).locator("#ao-compact-view")).toBeHidden({ timeout });
  await expect(ao(page).locator("#ao-notice")).not.toHaveText("", { timeout: 10000 });
  return (await ao(page).locator("#ao-notice").textContent()) ?? "";
}
// Pause demandée par la page (crochet « ao-stop » de panel.js) dès que la page
// a vu `n` recherches complètes.
export async function pauseApres(page: Page, n: number, cle = "__typed") {
  await page.evaluate(([n, cle]) => {
    const tick = () => { if (((window as any)[cle as string] || []).length >= (n as number)) document.dispatchEvent(new Event("ao-stop")); else requestAnimationFrame(tick); };
    requestAnimationFrame(tick);
  }, [n, cle] as const);
}

export type Etat = { typed: string[]; typedCc: string[]; keys: string[]; violations: string[]; plusClics: number; champA: string; pastilles: number; corps: string; objet: string; recherche: string; rendus: [string, number][] };
export async function etat(page: Page): Promise<Etat> {
  return page.evaluate(() => {
    const w = window as any;
    const libre = (el: Element | null) => {
      if (!el) return "";
      if ((el as HTMLInputElement).value !== undefined && (el.tagName === "INPUT" || el.tagName === "TEXTAREA")) return (el as HTMLInputElement).value;
      let s = ""; el.childNodes.forEach((n: any) => { if (n.nodeType === 3) s += n.data; else if (!(n.classList && (n.classList.contains("_Entity") || n.hasAttribute("data-hovercard-id")))) s += n.textContent; });
      return s.replace(/​/g, "").trim();
    };
    const a = document.querySelector("#champ-a") || document.querySelector('div[name="to"] input') || document.querySelector('input[aria-label="To"]');
    return {
      typed: [...(w.__typed || [])], typedCc: [...(w.__typedCc || [])], keys: [...(w.__keys || [])], violations: [...(w.__violations || [])],
      plusClics: w.__plusClics || 0, champA: libre(a),
      pastilles: document.querySelectorAll('#champ-a ._Entity, div[name="to"] [data-hovercard-id]').length,
      corps: (document.querySelector("#corps") as HTMLElement | null)?.innerText ?? "",
      objet: (document.querySelector("#objet") as HTMLInputElement | null)?.value ?? "",
      recherche: (document.querySelector("#recherche") as HTMLInputElement | null)?.value ?? "",
      rendus: [...(w.__rendus || [])],
    };
  });
}

export const stockage = (worker: Worker) => worker.evaluate(() => chrome.storage.local.get(null)) as Promise<Record<string, any>>;
// Les adresses relevées par le passage (fichier .txt du moteur) et l'historique.
export async function releve(worker: Worker, source: "outlook" | "gmail" = "outlook") {
  const all = await stockage(worker);
  const cle = Object.keys(all).find((k) => k.includes("adresses-" + source + "-") && k.endsWith(".txt"));
  const txt = String((cle && all[cle]) || "").split("\n").map((l) => l.trim().toLowerCase()).filter(Boolean).sort();
  const base = String(all["base-" + source] || "").split("\n").map((l) => l.trim()).filter(Boolean).sort();
  const annuaire = String(all["mega-annuaire"] || "").split("\n").map((l) => l.trim()).filter(Boolean).sort();
  return { txt, base, annuaire, journal: all["reprise-" + source], actif: all["actif-" + source], all };
}
export const interdites = (keys: string[]) => keys.filter((k) => ["Enter", "Tab", ",", ";"].includes(k));

// Planificateur et lecture de lignes de Smart Mail, exécutés dans Node (mêmes
// fichiers que l'extension, chargés tels quels) : la séquence attendue.
const g: any = {};
new Function("globalThis", "module", "require", lire(join(EXT, "adresses", "core", "planif.js")))(g, undefined, undefined);
new Function("globalThis", "module", "require", lire(join(EXT, "adresses", "core", "moteur.js")))(g, undefined, undefined);
export const Planif = g.Planif;
export const Moteur = g.Moteur;
export function simuler(ask: (q: string) => { name?: string; nom?: string; mail: string }[], prefixe = "") {
  const plan = new Planif(prefixe);
  const seq: string[] = [], distinct = new Set<string>();
  for (let garde = 0; garde < 5000; garde++) {
    const q = plan.next();
    if (q === null) break;
    seq.push(q);
    const list = ask(q).map((x) => ({ name: x.name ?? x.nom ?? "", mail: x.mail }));
    plan.report(q, list.map((x) => x.name), list.map((x) => x.mail), list.map(() => true));
    list.forEach((x) => distinct.add(x.mail.toLowerCase()));
  }
  return { seq, distinct: [...distinct].sort() };
}
