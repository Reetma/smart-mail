// Banc commun : Smart Mail (extension/) chargée dans un Chromium headless au
// profil temporaire distinct, fausse page « nouvel Outlook » servie par
// context.route sous https://outlook.cloud.microsoft/ ; TOUT le reste est bloqué
// (aucune requête réelle, aucun vrai mail).
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import type { Cfg, Plan, Releve, Envoye } from "./generateur";

const ICI = dirname(fileURLToPath(import.meta.url));
export const EXT = join(ICI, "..", "..", "..", "extension");
const JS = readFileSync(join(ICI, "fausse-outlook-fuzz.js"), "utf8");
export const URL_OUTLOOK = "https://outlook.cloud.microsoft/mail/0/";

function pageHtml(cfg: Cfg): string {
  return `<!doctype html>
<html lang="fr" dir="ltr" data-pp-test="rapide">
<head><meta charset="utf-8"><title>Courrier - Outlook</title></head>
<body>
<script>window.__CFG = ${JSON.stringify(cfg).replace(/</g, "\\u003c")};</script>
<script src="/fuzz/fausse-outlook-fuzz.js"></script>
</body>
</html>`;
}

export type Banc = { context: BrowserContext; sw: Worker; page: Page; servir: (cfg: Cfg) => void; requetesBloquees: string[] };

export const test = base.extend<{ banc: Banc }>({
  // eslint-disable-next-line no-empty-pattern
  banc: async ({}, use) => {
    const dir = mkdtempSync(join(tmpdir(), "audit-integrite-envoi-fuzz-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: 1400, height: 900 },
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    let cfgCourante: Cfg | null = null;
    const requetesBloquees: string[] = [];
    // Tout ce qui n'est pas la fausse page est refusé (enregistré d'abord : les routes
    // plus récentes passent avant).
    await context.route("**/*", (route) => { requetesBloquees.push(route.request().url()); return route.abort(); });
    await context.route(/^https:\/\/outlook\.cloud\.microsoft\//, (route) => {
      const u = new URL(route.request().url());
      if (u.pathname === "/fuzz/fausse-outlook-fuzz.js") return route.fulfill({ status: 200, contentType: "text/javascript; charset=utf-8", body: JS });
      if (u.pathname.startsWith("/mail/") && cfgCourante) return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: pageHtml(cfgCourante) });
      requetesBloquees.push(u.href);
      return route.abort();
    });
    let [sw] = context.serviceWorkers();
    if (!sw) sw = await context.waitForEvent("serviceworker");
    const page = context.pages()[0] ?? (await context.newPage());
    page.on("dialog", (d) => void d.dismiss());
    await use({ context, sw, page, servir: (cfg) => { cfgCourante = cfg; }, requetesBloquees });
    await context.close();
    try { rmSync(dir, { recursive: true, force: true }); } catch { /* profil temporaire */ }
  },
});
export { expect };

const pp = (page: Page) => page.locator("[data-pp-panneau]");

export async function ouvrirOutlook(b: Banc, cfg: Cfg): Promise<void> {
  b.servir(cfg);
  await b.sw.evaluate(() => chrome.storage.local.set({ "mega-onglet": "envoi" }));
  await b.page.goto(URL_OUTLOOK);
  await b.page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ timeout: 20000 });
  await b.page.locator('html[data-mega-pret="1"]').waitFor({ timeout: 20000 });
  await expect(pp(b.page).locator("#pp-go")).toBeVisible();
}

// Remplit le panneau comme l'utilisateur (destinataires collés, objet, message collé, réglages).
export async function preparer(page: Page, plan: Plan): Promise<string> {
  const p = pp(page);
  await p.locator("#pp-dest").fill(plan.saisie);
  await p.locator("#pp-objet").fill(plan.objet);
  await page.evaluate((h) => {
    const ed = document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-texte")!;
    ed.innerHTML = "";
    ed.focus();
    const dt = new DataTransfer();
    dt.setData("text/html", h);
    dt.setData("text/plain", "texte brut");
    ed.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, plan.html);
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  await p.locator("#pp-max").fill("100000");
  if (plan.cci) {
    await p.locator("#pp-cci").check();
    await p.locator("#pp-paquet").fill(String(plan.paquet));
    await p.locator("#pp-paquet").blur();
  }
  if (plan.importance) await p.locator("#pp-importance").check();
  if (plan.sansSignature) await p.locator("#pp-sans-signature").check();
  await page.waitForTimeout(350); // compte des adresses (mis à jour 250 ms après la saisie)
  return page.evaluate(() => {
    const o = document.querySelector("[data-pp-panneau]")!.shadowRoot!;
    return ["pp-nb", "pp-dest-info", "pp-notice"].map((id) => o.getElementById(id)?.textContent || "").join(" | ");
  });
}

// Lance l'envoi et attend la fin (tout envoyé) ou l'arrêt (message affiché).
export async function lancer(page: Page, delaiMs = 60000): Promise<boolean> {
  const p = pp(page);
  await p.locator("#pp-go").click();
  const t0 = Date.now();
  for (;;) {
    const etat = await page.evaluate(() => {
      const o = document.querySelector("[data-pp-panneau]")!.shadowRoot!;
      return { live: !o.getElementById("pp-live")!.hidden, notice: (o.getElementById("pp-notice")!.textContent || "").trim() };
    });
    if (!etat.live && etat.notice) return true;
    if (Date.now() - t0 > delaiMs) return false;
    await page.waitForTimeout(100);
  }
}

export async function relever(b: Banc, fini: boolean, infos = ""): Promise<Releve> {
  const page = b.page;
  const sto = await b.sw.evaluate(() => chrome.storage.local.get(["pp-campagne", "pp-liste"]));
  const c = sto["pp-campagne"] as { html: string; objet: string; resultats: Record<string, { s: string }> } | undefined;
  const liste = (sto["pp-liste"] as string[] | undefined) ?? [];
  const statuts: Record<string, string> = {};
  if (c) Object.entries(c.resultats).forEach(([k, v]) => { const a = liste[+k]; if (a) statuts[a.toLowerCase()] = v.s; });
  const vu = await page.evaluate((html) => {
    const lire = (h: string) => {
      const d = new DOMParser().parseFromString("<div id='r'>" + (h || "") + "</div>", "text/html").getElementById("r")!;
      return { texte: d.textContent || "", liens: [...d.querySelectorAll("a[href]")].map((a) => ({ href: a.getAttribute("href") || "", texte: a.textContent || "" })) };
    };
    const f = (window as unknown as { __faux: { envoyes: Array<Record<string, unknown>>; journal: Array<{ type: string; detail: unknown }> } }).__faux;
    const envoyes = f.envoyes.map((m) => {
      const sigs = (m.signatures as string[]).map(lire), restes = (m.restes as string[]).map(lire);
      return { ...m, signaturesTexte: sigs.map((s) => s.texte), signaturesLiens: sigs.flatMap((s) => s.liens.map((l) => l.href)), restesTexte: restes.map((s) => s.texte) };
    });
    const o = document.querySelector("[data-pp-panneau]")!.shadowRoot!;
    const n = o.getElementById("pp-notice")!;
    return {
      envoyes, journal: f.journal, notice: (n.textContent || "").trim(),
      noticeType: n.classList.contains("warn") ? "warn" : n.classList.contains("ok") ? "ok" : "",
      prevu: html ? lire(html) : null,
    };
  }, c ? c.html : "");
  return { envoyes: vu.envoyes as unknown as Envoye[], journal: vu.journal as Releve["journal"], notice: vu.notice, noticeType: vu.noticeType, statuts, liste, prevu: vu.prevu, objetCampagne: c ? c.objet : null, fini, infos };
}

// Une exécution complète : page, panneau, envoi, relevé.
export async function executer(b: Banc, cfg: Cfg, plan: Plan): Promise<Releve> {
  await ouvrirOutlook(b, cfg);
  const infos = await preparer(b.page, plan);
  const fini = await lancer(b.page);
  return relever(b, fini, infos);
}

export function rapport(plan: Plan, r: Releve, divergences: string[]): string {
  const j = r.journal.filter((x) => !["etape", "touche"].includes(x.type)).slice(-40).map((x) => x.type + (x.detail ? " " + JSON.stringify(x.detail) : "")).join("\n    ");
  return [
    `graine ${plan.graine} — variantes : ${plan.resume.join(", ")}`,
    `prévu : ${plan.destinataires.join(", ")}`,
    `objet prévu : ${plan.objet}`,
    `envoyés (${r.envoyes.length}) : ${r.envoyes.map((m) => `[À:${m.a.join(",")} Cc:${m.cc.join(",")} Cci:${m.cci.join(",")}]`).join(" ")}`,
    `panneau : « ${r.notice} » (${r.noticeType})`,
    `DIVERGENCES :\n  - ${divergences.join("\n  - ")}`,
    `journal (fin) :\n    ${j}`,
  ].join("\n");
}
