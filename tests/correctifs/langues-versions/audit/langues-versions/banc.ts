// Banc commun de l'audit « Langues et versions d'Outlook » (Smart Mail).
//
// - Chromium headless, extension réelle (extension/) chargée dans un profil
//   temporaire NEUF pour chaque test.
// - AUCUNE requête réelle : toutes les adresses Outlook (outlook.office.com,
//   outlook.office365.com, outlook.live.com, outlook.cloud.microsoft) sont servies
//   par context.route (fausse page configurable outlook-variante.html) ; tout le
//   reste est refusé (route.abort) et, par sécurité, la résolution DNS est coupée
//   (--host-resolver-rules) : rien ne peut sortir de la machine.
// - La fausse page journalise tout (window.__journal) et enregistre les mails
//   « envoyés » (data-envoyes, et localStorage « __envoyesGlobal » toutes fenêtres).
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker, type Locator } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "..", "..", "extension");
const PAGE = readFileSync(join(ICI, "outlook-variante.html"), "utf8");
const SCRIPTS: Record<string, string> = {
  "libelles.js": readFileSync(join(ICI, "libelles.js"), "utf8"),
  "annuaire.js": readFileSync(join(ICI, "annuaire.js"), "utf8"),
  "outlook-variante.js": readFileSync(join(ICI, "outlook-variante.js"), "utf8"),
};
// Le carnet factice (annuaire.js, script classique) chargé côté test.
export const ANNUAIRE: { adresses: string[]; CAP: number } = (() => {
  const m = { exports: {} as { adresses: string[]; CAP: number } };
  new Function("module", SCRIPTS["annuaire.js"])(m);
  return m.exports;
})();
const HOTES = /^https:\/\/(outlook\.office\.com|outlook\.office365\.com|outlook\.live\.com|outlook\.cloud\.microsoft)\//;

export type Options = { largeur?: number; hauteur?: number; onglet?: "envoi" | "adresses" };

export const test = base.extend<{ context: BrowserContext; worker: Worker; requetesBloquees: string[]; options: Options }>({
  options: [{}, { option: true }],
  requetesBloquees: async ({}, use) => { await use([]); },
  // eslint-disable-next-line no-empty-pattern
  context: async ({ options, requetesBloquees }, use) => {
    const dir = mkdtempSync(join(tmpdir(), "audit-langues-versions-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: options.largeur ?? 1400, height: options.hauteur ?? 900 },
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`, "--host-resolver-rules=MAP * ~NOTFOUND"],
    });
    // Les routes enregistrées en dernier passent en premier : refus par défaut, puis Outlook factice.
    await context.route("**/*", (route) => { requetesBloquees.push(route.request().url()); return route.abort(); });
    await context.route(HOTES, (route) => {
      const p = new URL(route.request().url()).pathname;
      if (p.startsWith("/__audit/")) {
        const nom = p.slice("/__audit/".length);
        if (SCRIPTS[nom] != null) return route.fulfill({ status: 200, contentType: "text/javascript; charset=utf-8", body: SCRIPTS[nom] });
        return route.fulfill({ status: 404, body: "" });
      }
      if (/\.(ico|png|svg|js|css|json)$/.test(p)) return route.fulfill({ status: 404, body: "" });
      return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: PAGE });
    });
    await use(context);
    await context.close();
    try { rmSync(dir, { recursive: true, force: true }); } catch { /* profil temporaire */ }
  },
  worker: async ({ context }, use) => {
    let [w] = context.serviceWorkers();
    if (!w) w = await context.waitForEvent("serviceworker");
    await use(w);
  },
});
export { expect };
export type { Page, BrowserContext, Worker } from "@playwright/test";

// ------------------------------------------------------------ navigation
export async function ouvrir(page: Page, url: string): Promise<void> {
  await page.goto(url);
  await page.bringToFront();
  await page.locator('html[data-faux-pret="1"]').waitFor({ timeout: 15000 });
  await page.locator('html[data-mega-pret="1"]').waitFor({ timeout: 30000 });
}
export async function premierePage(context: BrowserContext): Promise<Page> {
  return context.pages()[0] ?? (await context.newPage());
}
export async function choisirOnglet(worker: Worker, onglet: "envoi" | "adresses"): Promise<void> {
  await worker.evaluate((o) => chrome.storage.local.set({ "mega-onglet": o }), onglet);
}
export function url(params: Record<string, string | number | undefined>, hote = "outlook.office.com", chemin = "/mail/"): string {
  const q = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) if (v !== undefined && v !== "") q.set(k, String(v));
  return `https://${hote}${chemin}?${q.toString()}`;
}

// ------------------------------------------------------------ panneaux
export const pp = (page: Page): Locator => page.locator("[data-pp-panneau]");
export const ao = (page: Page): Locator => page.locator('[data-testid="ao-panel"]');
export const pastilleReleve = (page: Page): Locator => ao(page).locator('.mega-tab[data-onglet="adresses"] .pastille');

// ------------------------------------------------------------ état de la fausse page
export type Mail = { a: string[]; cc: string[]; cci: string[]; objet: string; corps: string; html: string; importance: string; fenetre: string };
export type Journal = {
  clics: Array<{ lib: string; id: string; cmd: string; zone: string; confiance: boolean }>;
  saisies: Array<{ champ: string; type: string; data: string | null; confiance: boolean }>;
  touches: Array<{ champ: string; key: string }>;
  carnet: number; teamsEnvoye: number; teamsMessages: string[]; suggestionChoisie: number; menus: string[];
  textes: Array<{ champ: string; libre: string }>;
};
export type Champ = { adresses: string[]; libre: string; pastilles: number } | null;
export type Etat = {
  redaction: { id: string; a: Champ; cc: Champ; cci: Champ; objet: string; corpsTexte: string; corpsHtml: string; importance: string } | null;
  dialogue: boolean; menus: number; onglet: string; actif: string; recherche: string; teams: string;
};
export const envoyes = async (page: Page): Promise<Mail[]> => JSON.parse((await page.locator("html").getAttribute("data-envoyes")) ?? "[]");
export const envoyesGlobal = async (page: Page): Promise<Mail[]> => page.evaluate(() => JSON.parse(localStorage.getItem("__envoyesGlobal") || "[]"));
export const journal = async (page: Page): Promise<Journal> => page.evaluate(() => JSON.parse(JSON.stringify((window as any).__journal)));
export const etat = async (page: Page): Promise<Etat> => page.evaluate(() => (window as any).__etat());
export const razJournal = async (page: Page): Promise<void> => { await page.evaluate(() => (window as any).__razJournal()); };
export const sansInvisibles = (s: string | null | undefined): string => String(s ?? "").replace(/[\s​-‍⁠﻿]/g, "");

// Clics de l'EXTENSION (événements non « isTrusted ») sur une commande qu'elle ne
// doit jamais toucher (règle E6) : carnet d'adresses, Abandonner, Isoler, Fermer,
// menu d'envoi, Envoyer de Teams, liste des messages, suggestion (= validation).
const CMD_INTERDITES = ["carnet", "carnet-a", "carnet-cc", "carnet-cci", "abandonner", "isoler", "fermer", "envoyer-plus", "teams-envoyer", "arriere-plan", "nouveau-menu", "teams"];
const ZONES_INTERDITES = ["liste", "teams", "dialogue", "entete", "nav", "suggestions-option"];
export function clicsInterdits(j: Journal): Journal["clics"] {
  return j.clics.filter((c) => !c.confiance && (CMD_INTERDITES.includes(c.cmd) || ZONES_INTERDITES.includes(c.zone)));
}
// Champs où l'extension a tapé ou effacé (événements input), hors de la liste permise.
export function saisiesHors(j: Journal, permis: string[]): Journal["saisies"] {
  return j.saisies.filter((s) => !permis.includes(s.champ));
}
// Touches qui valideraient un destinataire (règle R1).
export function touchesValidation(j: Journal): Journal["touches"] {
  return j.touches.filter((t) => ["Enter", "Tab", ";", ","].includes(t.key));
}

// ------------------------------------------------------------ envoi (onglet ② Envoyer)
export type Envoi = { dest: string[]; objet?: string; texte?: string; cci?: boolean; paquet?: number; importance?: boolean; html?: string };
export async function remplirEnvoi(page: Page, e: Envoi): Promise<void> {
  const p = pp(page);
  await p.locator("#pp-dest").fill(e.dest.join("\n"));
  await p.locator("#pp-objet").fill(e.objet ?? "Réunion de rentrée");
  if (e.html != null) {
    await page.evaluate((h) => {
      const ed = document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-texte")!;
      ed.innerHTML = ""; ed.focus();
      const dt = new DataTransfer(); dt.setData("text/html", h); dt.setData("text/plain", "texte");
      ed.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    }, e.html);
  } else await p.locator("#pp-texte").fill(e.texte ?? "Bonjour,\nRendez-vous lundi à 9 h.\nÀ bientôt");
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  if (e.cci) { await p.locator("#pp-cci").check(); if (e.paquet) await p.locator("#pp-paquet").fill(String(e.paquet)); }
  if (e.importance) await p.locator("#pp-importance").check();
}
export async function lancerEnvoi(page: Page, e: Envoi): Promise<void> {
  await remplirEnvoi(page, e);
  await razJournal(page);
  await pp(page).locator("#pp-go").click();
}
// Fin de l'envoi : le panneau revient à la préparation et affiche son bilan.
export async function finEnvoi(page: Page, delai = 60000): Promise<string> {
  await expect.poll(async () => ((await pp(page).locator("#pp-notice").textContent()) ?? "").trim(), { timeout: delai }).not.toBe("");
  await expect(pp(page).locator("#pp-live")).toBeHidden({ timeout: delai });
  return ((await pp(page).locator("#pp-notice").textContent()) ?? "").trim();
}

// ------------------------------------------------------------ relevé (onglet ① Relever)
export async function demarrerReleve(page: Page): Promise<void> {
  await razJournal(page);
  await ao(page).locator("#ao-run").click();
}
export async function adressesReleve(page: Page): Promise<string[]> {
  return (await ao(page).locator("#ao-rows .res .m").allTextContents()).map((s) => s.trim().toLowerCase());
}
export async function noticeReleve(page: Page): Promise<string> {
  return ((await ao(page).locator("#ao-notice").textContent()) ?? "").trim();
}
// Met le relevé en pause (s'il tourne) et attend qu'il ait rendu la main.
export async function pauseReleve(page: Page): Promise<void> {
  if (await pastilleReleve(page).isVisible()) {
    await ao(page).locator("#ao-stop").click().catch(() => {});
    await expect(pastilleReleve(page)).toBeHidden({ timeout: 20000 });
  }
}
// Attend qu'une condition soit vraie, ou au plus « ms » : rend la dernière valeur.
export async function attendreJusqua<T>(fn: () => Promise<T>, ok: (v: T) => boolean, ms: number, pas = 250): Promise<T> {
  const fin = Date.now() + ms;
  let v = await fn();
  while (!ok(v) && Date.now() < fin) { await new Promise((r) => setTimeout(r, pas)); v = await fn(); }
  return v;
}
