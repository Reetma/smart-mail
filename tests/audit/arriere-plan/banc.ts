// Banc « arrière-plan et cycle de vie » pour Smart Mail (extension/).
// - Chromium headless (channel "chromium"), profil temporaire distinct par test,
//   réouvrable (même dossier) pour simuler fermeture/réouverture du navigateur.
// - AUCUNE requête réelle : context.route sert la fausse page Outlook
//   (outlook-reel.html) sous https://outlook.office.com et bloque tout le reste.
//   L'envoi d'un mail = requête POST /owa/service.svc?action=SendItem, enregistrée
//   ICI (côté Node) : le journal survit aux rechargements et aux redémarrages.
// - CDP : monde isolé de l'extension (contexte « Smart Mail »), gel de la page
//   (Debugger.pause : plus aucune tâche ne tourne dans la page ni dans le monde de
//   l'extension, le service worker continue, comme un onglet gelé), arrêt du
//   service worker (ServiceWorker.stopAllWorkers).
// - Émulation de l'onglet caché (emulation.js) dans les DEUX mondes.
import { chromium, expect, type BrowserContext, type CDPSession, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "extension");
const PAGE = readFileSync(join(ICI, "outlook-reel.html"), "utf8");
const EMULATION = readFileSync(join(ICI, "emulation.js"), "utf8");
export const URL_OUTLOOK = "https://outlook.office.com/mail/";

export type Envoi = { t: number; onglet: string; a: string[]; cci: string[]; objet: string; corps: string; html: string; importance: string };
export type Evt = { t: number; type: string; onglet: string; [k: string]: unknown };
export type CfgPage = {
  rapide?: boolean; suggestions?: string | null; pastille?: string; redactionRaf?: boolean; recharger?: string | null;
  autoRecharge?: number; redirection?: number; restaurerBrouillon?: boolean; bloque?: boolean; fermetureLente?: number;
};

export const attendre = (ms: number) => new Promise((r) => setTimeout(r, ms));

export class Banc {
  dir = mkdtempSync(join(tmpdir(), "sm-arriere-plan-"));
  context!: BrowserContext;
  cfg: CfgPage = { rapide: true };
  envois: Envoi[] = [];
  evts: Evt[] = [];
  surEnvoi: ((e: Envoi) => void | Promise<void>) | null = null;
  graine = 1;

  static async ouvrir(cfg: CfgPage = {}, graine = 1): Promise<Banc> {
    const b = new Banc();
    b.cfg = { rapide: true, ...cfg };
    b.graine = graine;
    await b.lancer();
    return b;
  }

  async lancer(): Promise<void> {
    this.context = await chromium.launchPersistentContext(this.dir, {
      channel: "chromium",
      headless: true,
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    // Monde de la page : émulation installée avant tout script de la page.
    await this.context.addInitScript({ content: EMULATION + `\n;__emuler(null, "page");` });
    await this.context.route("**/*", async (route) => {
      const req = route.request();
      let u: URL;
      try { u = new URL(req.url()); } catch { return route.abort(); }
      if (u.hostname === "outlook.office.com" || u.hostname === "outlook.cloud.microsoft") {
        if (u.pathname.startsWith("/owa/service.svc") && req.method() === "POST") {
          const e = { ...(JSON.parse(req.postData() || "{}") as Envoi), t: Date.now() };
          this.envois.push(e);
          if (this.surEnvoi) await this.surEnvoi(e);
          return route.fulfill({ status: 200, contentType: "application/json", body: '{"ok":true}' }).catch(() => {});
        }
        if (u.pathname === "/__evt") {
          try { this.evts.push({ ...(JSON.parse(req.postData() || "{}") as Evt), t: Date.now() }); } catch { /* */ }
          return route.fulfill({ status: 204, body: "" }).catch(() => {});
        }
        return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: this.html() }).catch(() => {});
      }
      if (u.hostname === "login.microsoftonline.com") {
        // Session expirée : page neutre, sans formulaire (fausse page de test).
        return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: "<!doctype html><meta charset=utf-8><title>Connexion</title><p>Session expirée (fausse page de test)</p>" }).catch(() => {});
      }
      return route.abort().catch(() => {});
    });
  }

  html(): string {
    const { rapide, ...reste } = this.cfg;
    return PAGE.replace("<!--ATTR-->", rapide ? ' data-pp-test="rapide"' : "").replace("<!--CONFIG-->", `<script>window.__CFG = ${JSON.stringify(reste)};</script>`);
  }

  // Fermeture « propre » du navigateur (comme Cmd+Q), puis réouverture sur le même profil.
  async fermerNavigateur(): Promise<void> { await this.context.close().catch(() => {}); }
  async rouvrir(): Promise<void> { await this.lancer(); }

  async fin(): Promise<void> {
    await this.context.close().catch(() => {});
    try { rmSync(this.dir, { recursive: true, force: true }); } catch { /* */ }
  }

  // « ↻ Recharger » de chrome://extensions : l'extension redémarre (nouvelle version
  // active), les scripts déjà injectés dans les onglets ouverts deviennent orphelins.
  // (Chargée par --load-extension, Chromium la laisse désactivée après
  // runtime.reload() : on la réactive comme le ferait le bouton.)
  async rechargerExtension(): Promise<void> {
    const w = await this.sw();
    const id = new URL(w.url()).host;
    await w.evaluate(() => chrome.runtime.reload()).catch(() => {});
    await attendre(800);
    const p = await this.context.newPage();
    await p.goto("chrome://extensions/");
    await expect.poll(() => p.evaluate((x) => new Promise<string>((r) => (globalThis as any).chrome.management.setEnabled(x, true, () => r((globalThis as any).chrome.runtime.lastError ? "erreur" : "ok"))), id), { timeout: 10000 }).toBe("ok");
    await p.close();
    await this.sw(15000);
  }
  // Interrupteur de chrome://extensions.
  idExtension = "";
  async activerExtension(actif: boolean): Promise<void> {
    if (!this.idExtension) this.idExtension = new URL((await this.sw()).url()).host;
    const p = await this.context.newPage();
    await p.goto("chrome://extensions/");
    await expect.poll(() => p.evaluate(([x, a]) => new Promise<string>((r) => (globalThis as any).chrome.management.setEnabled(x, a, () => r((globalThis as any).chrome.runtime.lastError ? "erreur" : "ok"))), [this.idExtension, actif] as [string, boolean]), { timeout: 10000 }).toBe("ok");
    await p.close();
    if (actif) await this.sw(15000);
  }

  async sw(timeout = 10000): Promise<Worker> {
    const t0 = Date.now();
    for (;;) {
      const w = this.context.serviceWorkers().find((x) => x.url().startsWith("chrome-extension://"));
      if (w) return w;
      if (Date.now() - t0 > timeout) throw new Error("service worker introuvable");
      await this.context.waitForEvent("serviceworker", { timeout: 1000 }).catch(() => {});
    }
  }
  async stockage(cles: string[] | null = null): Promise<Record<string, unknown>> {
    const w = await this.sw();
    return w.evaluate((c) => chrome.storage.local.get(c), cles);
  }
  async ecrireStockage(o: Record<string, unknown>): Promise<void> {
    const w = await this.sw();
    await w.evaluate((x) => chrome.storage.local.set(x), o);
  }

  // Ouvre Outlook dans un onglet (nouveau, ou le premier vide) et attend les deux panneaux.
  async onglet(opts: { onglet?: "envoi" | "adresses"; nouveau?: boolean } = {}): Promise<Onglet> {
    if (opts.onglet) await this.ecrireStockage({ "mega-onglet": opts.onglet });
    const vide = this.context.pages().find((p) => p.url() === "about:blank");
    const page = !opts.nouveau && vide ? vide : await this.context.newPage();
    const o = new Onglet(this, page);
    await o.brancher();
    await page.goto(URL_OUTLOOK);
    await o.pret();
    return o;
  }

  envoisA(adresse: string): Envoi[] { return this.envois.filter((e) => e.a.includes(adresse) || e.cci.includes(adresse)); }
  // Destinataires servis plus d'une fois (règle E2).
  doublons(): string[] {
    const n = new Map<string, number>();
    for (const e of this.envois) for (const x of [...e.a, ...e.cci]) n.set(x, (n.get(x) ?? 0) + 1);
    return [...n].filter(([, k]) => k > 1).map(([x, k]) => `${x} ×${k}`);
  }
}

export class Onglet {
  cdp!: CDPSession;
  iso: number | null = null;
  cadre: string | null = null;
  gele = false;
  constructor(public banc: Banc, public page: Page) {}

  async brancher(): Promise<void> {
    this.cdp = await this.banc.context.newCDPSession(this.page);
    this.cdp.on("Runtime.executionContextCreated", (e: { context: { id: number; name: string; auxData?: { frameId?: string; type?: string } } }) => {
      const c = e.context;
      if (c.name === "Smart Mail" && c.auxData?.type === "isolated" && (!this.cadre || c.auxData.frameId === this.cadre)) this.iso = c.id;
    });
    this.cdp.on("Runtime.executionContextsCleared", () => { this.iso = null; });
    this.cdp.on("Runtime.executionContextDestroyed", (e: { executionContextId: number }) => { if (e.executionContextId === this.iso) this.iso = null; });
    await this.cdp.send("Page.enable");
    const arbre = (await this.cdp.send("Page.getFrameTree")) as { frameTree: { frame: { id: string } } };
    this.cadre = arbre.frameTree.frame.id;
    await this.cdp.send("Runtime.enable");
  }

  // Panneaux prêts ; émulation installée dans le monde isolé de l'extension.
  async pret(timeout = 30000): Promise<void> {
    await this.page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached", timeout });
    await this.page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached", timeout });
    await expect.poll(async () => { try { return await this.evalIso<string>("typeof PP === 'object' && !!PP.Outlook ? 'ok' : 'non'"); } catch { return "absent"; } }, { timeout: 15000 }).toBe("ok");
    await this.evalIso(EMULATION + `\n;__emuler(${this.banc.graine}, "extension");`);
  }
  async recharger(): Promise<void> {
    this.iso = null;
    await this.page.reload();
    await this.pret();
  }
  // Après un rechargement venu de la page elle-même (Outlook qui se recharge) :
  // attend le chargement suivant le n-ième (window.__chargements, tenu par la page).
  async apresRechargement(n = 1): Promise<void> {
    await expect.poll(() => this.page.evaluate(() => (window as any).__chargements ?? 0).catch(() => 0), { timeout: 60000 }).toBeGreaterThan(n);
    await this.pret();
  }
  // Chrome arrête le service worker de l'extension (MV3 : inactivité, mémoire…).
  async arreterSW(): Promise<void> {
    await this.cdp.send("ServiceWorker.enable").catch(() => {});
    await this.cdp.send("ServiceWorker.stopAllWorkers");
  }

  async evalIso<T = unknown>(expression: string): Promise<T> {
    if (this.iso == null) throw new Error("contexte de l'extension introuvable");
    const r = (await this.cdp.send("Runtime.evaluate", { contextId: this.iso, expression, returnByValue: true, awaitPromise: true })) as { result: { value: T }; exceptionDetails?: { text: string; exception?: { description?: string } } };
    if (r.exceptionDetails) throw new Error("evalIso : " + (r.exceptionDetails.exception?.description || r.exceptionDetails.text));
    return r.result.value;
  }

  // Onglet caché / remis devant. mode "seconde" : 1 réveil/s ; "minute" : caché depuis plus de 5 min.
  async cacher(mode: "seconde" | "minute" = "seconde"): Promise<void> {
    const o = JSON.stringify({ cache: true, focus: false, mode });
    await this.evalIso(`__AP.etat(${o}, false)`);
    await this.page.evaluate((x) => (window as any).__AP.etat(JSON.parse(x), true), o);
  }
  // Fenêtre Chrome derrière une autre application : onglet visible mais sans le focus.
  async sansFocus(): Promise<void> {
    const o = JSON.stringify({ cache: false, focus: false, mode: "aucun" });
    await this.evalIso(`__AP.etat(${o}, false)`);
    await this.page.evaluate((x) => (window as any).__AP.etat(JSON.parse(x), true), o);
  }
  async montrer(): Promise<void> {
    const o = JSON.stringify({ cache: false, focus: true, mode: "aucun" });
    await this.evalIso(`__AP.etat(${o}, false)`);
    await this.page.evaluate((x) => (window as any).__AP.etat(JSON.parse(x), true), o);
  }
  // Saut de l'heure murale (veille du Mac), dans les deux mondes.
  async decaler(ms: number): Promise<void> {
    const o = JSON.stringify({ decalage: ms });
    if (this.gele) throw new Error("decaler() : poser le saut d'horloge avant geler()");
    await this.evalIso(`__AP.etat(${o}, false)`);
    await this.page.evaluate((x) => (window as any).__AP.etat(JSON.parse(x), false), o);
  }

  // Gel de l'onglet : plus rien ne s'exécute dans la page ni dans le monde de
  // l'extension (le service worker, lui, continue). Ne pas utiliser this.page
  // (evaluate, locator) tant que l'onglet est gelé.
  async geler(): Promise<void> {
    await this.cdp.send("Debugger.enable");
    const vu = new Promise<void>((r) => this.cdp.once("Debugger.paused", () => r()));
    await this.cdp.send("Debugger.pause");
    await Promise.race([vu, attendre(5000)]);
    this.gele = true;
  }
  async degeler(): Promise<void> {
    await this.cdp.send("Debugger.resume").catch(() => {});
    await this.cdp.send("Debugger.disable").catch(() => {});
    this.gele = false;
  }

  // ---------------------------------------------------------------- panneaux
  pp() { return this.page.locator("[data-pp-panneau]"); }
  ao() { return this.page.locator('[data-testid="ao-panel"]'); }
  async ongletEnvoi(): Promise<void> {
    if (await this.pp().isHidden()) await this.ao().locator('.mega-tab[data-onglet="envoi"]').click();
    await expect(this.pp()).toBeVisible();
  }
  async ongletAdresses(): Promise<void> {
    if (await this.ao().isHidden()) await this.pp().locator('.mega-tab[data-onglet="adresses"]').click();
    await expect(this.ao()).toBeVisible();
  }
  async preparerEnvoi(o: { dest: string[]; objet?: string; texte?: string; pmin?: number; pmax?: number; max?: number }): Promise<void> {
    await this.ongletEnvoi();
    const p = this.pp();
    await p.locator("#pp-dest").fill(o.dest.join("\n"));
    await p.locator("#pp-objet").fill(o.objet ?? "Invitation");
    await p.locator("#pp-texte").fill(o.texte ?? "Bonjour,\nVoici l'invitation.\nÀ bientôt");
    await p.locator("#pp-pmin").fill(String(o.pmin ?? 0));
    await p.locator("#pp-pmax").fill(String(o.pmax ?? o.pmin ?? 0));
    if (o.max != null) await p.locator("#pp-max").fill(String(o.max));
    await expect(p.locator("#pp-nb")).toHaveText(`${o.dest.length} adresse${o.dest.length > 1 ? "s" : ""}`);
  }
  async lancer(): Promise<void> { await this.pp().locator("#pp-go").click(); }
  notice() { return this.pp().locator("#pp-notice"); }
  async enPause(): Promise<boolean> { return (await this.pp().getAttribute("data-pp-pause")) != null && (await this.pp().locator("#pp-live-etape").textContent() ?? "").startsWith("Pause"); }
  async redactionOuverte(): Promise<boolean> { return this.page.evaluate(() => !!document.getElementById("redaction")); }
  async contenuA(): Promise<string> { return this.page.evaluate(() => { const a = document.querySelector('[aria-label="À"][contenteditable]'); return a ? (a.textContent || "").replace(/​/g, "") : ""; }); }
}

// Campagne vue dans le stockage de l'extension.
export type Campagne = { curseur: number; total: number; resultats: Record<string, { s: string; d?: string; t?: number }> };
export async function campagne(b: Banc): Promise<Campagne | undefined> { return (await b.stockage(["pp-campagne"]))["pp-campagne"] as Campagne | undefined; }
export async function envoisCompte(b: Banc): Promise<number> { return (((await b.stockage(["pp-envois"]))["pp-envois"] as number[]) ?? []).length; }
export const liste = (n: number, p = "d") => Array.from({ length: n }, (_, i) => `${p}${i + 1}@exemple.fr`);
