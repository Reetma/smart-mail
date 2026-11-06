// Banc commun de l'audit « fidélité au vrai nouvel Outlook ». Smart Mail (extension/)
// est chargée telle quelle dans Chromium ; la fausse page outlook-reel.html (+ .js)
// est servie à https://outlook.cloud.microsoft/… par context.route ; TOUTE autre
// requête est bloquée (et notée). Rien ne sort, aucun vrai mail.
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "extension");
const HTML = readFileSync(join(ICI, "outlook-reel.html"), "utf8");
const JS = readFileSync(join(ICI, "outlook-reel.js"), "utf8");
const CONNEXION = '<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Se connecter à votre compte</title></head><body><h1>Connexion</h1><input type="email" aria-label="Adresse e-mail"></body></html>';

export const URL_OUTLOOK = "https://outlook.cloud.microsoft/mail/inbox";
export { expect };
export type Journal = Array<{ n: number; t: number; type: string; cible?: string; confiance?: boolean; touche?: string; genre?: string; data?: string | null; texte?: string; champ?: string; adresse?: string | null; source?: string; id?: string; ouverte?: boolean; raison?: string; [k: string]: unknown }>;
export type Envoye = { a: string[]; cc: string[]; cci: string[]; objet: string; corps: string; texte: string; html: string; importance: string; type: string; libres: string[]; t: number; depuisBoiteEnvoi?: boolean };
type Pastille = { adresse: string | null; nom: string | null; texte: string | null; resolu: boolean; source: string; affiche: string };
export type Etat = {
  composes: Array<{ id: string; type: string; visible: boolean; a: Pastille[]; cc: Pastille[]; cci: Pastille[]; libres: { A: string; CC: string; CCI: string }; objet: string | null; corps: string; html: string; importance: string; cciVisible: boolean }>;
  envoyes: number; boite: number; carnet: number; modales: string[]; menus: number; actif: string; onglet: string;
};

export const test = base.extend<{ context: BrowserContext; worker: Worker; requetes: string[] }>({
  // eslint-disable-next-line no-empty-pattern
  requetes: async ({}, use) => { await use([]); },
  context: async ({ requetes }, use) => {
    const dir = mkdtempSync(join(tmpdir(), "audit-outlook-fidele-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: 1400, height: 900 },
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    await context.route("**/*", (route) => {
      const u = new URL(route.request().url());
      if (u.hostname === "outlook.cloud.microsoft") {
        if (u.pathname === "/__fake/outlook-reel.js") return route.fulfill({ status: 200, contentType: "text/javascript; charset=utf-8", body: JS });
        return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: HTML });
      }
      if (u.hostname === "login.microsoftonline.com") return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: CONNEXION });
      requetes.push(u.href);
      return route.abort();
    });
    await use(context);
    await context.close();
    rmSync(dir, { recursive: true, force: true });
  },
  worker: async ({ context }, use) => {
    let [w] = context.serviceWorkers();
    if (!w) w = await context.waitForEvent("serviceworker");
    await use(w);
  },
});

export const pp = (page: Page) => page.locator("[data-pp-panneau]");
export const ao = (page: Page) => page.locator('[data-testid="ao-panel"]');
export const onglet = (page: Page, cote: "ao" | "pp", o: "adresses" | "envoi") => (cote === "ao" ? ao(page) : pp(page)).locator(`.mega-tab[data-onglet="${o}"]`);

// Compte les débuts / fins d'envoi et de relevé (vue « en cours » des deux panneaux),
// même quand un envoi accéléré se fait entre deux sondages.
async function instrumenter(page: Page) {
  await page.evaluate(() => {
    const w = window as unknown as Record<string, unknown>;
    const suivre = (cle: string, el: HTMLElement | null) => {
      const etat = { debuts: 0, fins: 0, visible: !!el && !el.hidden };
      w[cle] = etat;
      if (!el) return;
      new MutationObserver((recs) => {
        recs.forEach((r, i) => {
          const avant = r.oldValue === null; // visible avant cette mutation
          const apres = i + 1 < recs.length ? recs[i + 1]!.oldValue === null : !el.hidden;
          if (!avant && apres) etat.debuts++;
          if (avant && !apres) etat.fins++;
          etat.visible = apres;
        });
      }).observe(el, { attributes: true, attributeFilter: ["hidden"], attributeOldValue: true });
    };
    const hp = document.querySelector("[data-pp-panneau]") as HTMLElement & { shadowRoot: ShadowRoot };
    const ha = document.querySelector('[data-testid="ao-panel"]') as HTMLElement & { shadowRoot: ShadowRoot };
    suivre("__ppSuivi", hp.shadowRoot.getElementById("pp-live"));
    suivre("__aoSuivi", ha.shadowRoot.getElementById("ao-compact-view"));
  });
}

export type Options = { onglet?: "envoi" | "adresses"; vitesse?: "fast" | "normal" | "careful"; stockage?: Record<string, unknown>; url?: string };
export async function ouvrir(context: BrowserContext, worker: Worker, cfg: Record<string, unknown> = {}, o: Options = {}): Promise<Page> {
  await worker.evaluate((s) => chrome.storage.local.set(s), { "mega-onglet": o.onglet ?? "envoi", reglages: JSON.stringify({ Speed: o.vitesse ?? "fast" }), ...(o.stockage ?? {}) });
  await context.addInitScript((c) => { (window as unknown as { __OUTLOOK_CFG: unknown }).__OUTLOOK_CFG = c; }, cfg);
  const page = context.pages()[0] ?? (await context.newPage());
  await page.goto(o.url ?? URL_OUTLOOK);
  await page.bringToFront();
  await pret(page);
  return page;
}
export async function pret(page: Page) {
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await page.locator('[data-testid="ao-panel"][data-ao-ready="1"]').waitFor({ state: "attached", timeout: 30000 });
  await instrumenter(page);
}
export async function recharger(page: Page) { await page.reload(); await pret(page); }

export const journal = (page: Page) => page.evaluate(() => (window as unknown as { __journal: Journal }).__journal.slice()) as Promise<Journal>;
export const etat = (page: Page) => page.evaluate(() => (window as unknown as { __fausse: { etat: () => Etat } }).__fausse.etat()) as Promise<Etat>;
export const envoyes = async (page: Page) => JSON.parse((await page.locator("html").getAttribute("data-envoyes")) ?? "[]") as Envoye[];
export const suivi = (page: Page, cle: "__ppSuivi" | "__aoSuivi") => page.evaluate((k) => (window as unknown as Record<string, { debuts: number; fins: number; visible: boolean }>)[k], cle);
// Clics faits par l'EXTENSION (événements non « isTrusted ») sur la page Outlook.
export const clicsExtension = async (page: Page) => (await journal(page)).filter((e) => e.type === "clic" && !e.confiance && e.cible !== "PANNEAU").map((e) => e.cible!);

// ------------------------------------------------------------ envoi
export type Envoi = { dest: string[]; objet?: string; texte?: string; html?: string; pmin?: number; pmax?: number; max?: number; cci?: boolean; paquet?: number; importance?: boolean; sansSignature?: boolean };
export async function collerDansPanneau(page: Page, html: string) {
  await page.evaluate((h) => {
    const ed = (document.querySelector("[data-pp-panneau]") as HTMLElement & { shadowRoot: ShadowRoot }).shadowRoot.getElementById("pp-texte")!;
    ed.innerHTML = "";
    ed.focus();
    const dt = new DataTransfer();
    dt.setData("text/html", h);
    dt.setData("text/plain", "texte brut");
    ed.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, html);
}
export async function preparerEnvoi(page: Page, o: Envoi) {
  const p = pp(page);
  await expect(p).toBeVisible();
  await p.locator("#pp-dest").fill(o.dest.join("\n"));
  if (o.objet != null) await p.locator("#pp-objet").fill(o.objet);
  if (o.html != null) await collerDansPanneau(page, o.html);
  else if (o.texte != null) await p.locator("#pp-texte").fill(o.texte);
  await p.locator("#pp-pmin").fill(String(o.pmin ?? 0));
  await p.locator("#pp-pmax").fill(String(o.pmax ?? 0));
  if (o.max != null) await p.locator("#pp-max").fill(String(o.max));
  if (o.cci) { await p.locator("#pp-cci").check(); if (o.paquet != null) await p.locator("#pp-paquet").fill(String(o.paquet)); }
  if (o.importance) await p.locator("#pp-importance").check();
  if (o.sansSignature) await p.locator("#pp-sans-signature").check();
  await page.waitForTimeout(400); // brouillon du panneau enregistré
}
// Lance (ou reprend) l'envoi et attend qu'il s'arrête ; rend le message affiché.
export async function envoyer(page: Page, timeout = 60000): Promise<string> {
  const avant = (await suivi(page, "__ppSuivi")).fins;
  await pp(page).locator("#pp-go").click();
  await expect.poll(async () => (await suivi(page, "__ppSuivi")).fins, { timeout }).toBeGreaterThan(avant);
  return ((await pp(page).locator("#pp-notice").textContent()) ?? "").trim();
}
export async function lancerSansAttendre(page: Page) { await pp(page).locator("#pp-go").click(); }
export async function finEnvoi(page: Page, avant: number, timeout = 60000): Promise<string> {
  await expect.poll(async () => (await suivi(page, "__ppSuivi")).fins, { timeout }).toBeGreaterThan(avant);
  return ((await pp(page).locator("#pp-notice").textContent()) ?? "").trim();
}

// Texte sans aucun blanc (comparaison robuste des corps de mail).
export const compact = (s: string) => String(s ?? "").replace(/[\s ​﻿]+/g, "");
// Texte d'un HTML tel que rendu (sans blancs), calculé dans la page.
export const texteDe = (page: Page, html: string) => page.evaluate((h) => { const d = document.createElement("div"); d.innerHTML = h; return d.textContent ?? ""; }, html).then(compact);

// ------------------------------------------------------------ relevé
export async function demarrerReleve(page: Page) {
  const avant = (await suivi(page, "__aoSuivi")).debuts;
  await ao(page).locator("#ao-run").click();
  return avant;
}
export async function finReleve(page: Page, avant: number, timeout = 120000): Promise<string> {
  await expect.poll(async () => (await suivi(page, "__aoSuivi")).fins, { timeout }).toBeGreaterThan(avant);
  return ((await ao(page).locator("#ao-notice").textContent()) ?? "").trim();
}
export async function baseReleve(worker: Worker): Promise<string[]> {
  const o = await worker.evaluate(() => chrome.storage.local.get(["base-outlook", "mega-annuaire"]));
  return String(o["base-outlook"] || "").split("\n").filter(Boolean);
}
