// Audit « Fenêtres, politiques et messages d'Outlook selon l'établissement ».
//
// L'extension réelle (extension/) est chargée dans Chromium et pilote une fausse page
// du nouvel Outlook web (tests/audit/fenetres-politiques/outlook-etablissement.*),
// sur laquelle on active ce que les établissements configurent : étiquette de
// sensibilité obligatoire, avertissements destinataires externes, DLP, rappel de
// pièce jointe, « Annuler l'envoi », compléments (Letsignit…), destinataire non
// résolu, trop de destinataires, échec d'envoi, limitation, boîte pleine, session
// expirée, hors ligne, mise à jour d'Outlook, notifications.
//
// Règle d'or vérifiée partout : ne JAMAIS cliquer un bouton de ces fenêtres, ne
// jamais envoyer après une fenêtre inconnue, ne jamais compter comme envoyé ce qui
// ne l'est pas, ne jamais renvoyer ce qui est parti ; pendant un relevé : ne jamais
// taper dans une fenêtre, s'arrêter proprement (champ « À » vidé).
//
// Aucune requête réelle : outlook.office.com et login.microsoftonline.com sont servis
// par context.route, tout le reste est bloqué (et noté). Mode accéléré de la fausse
// page : <html data-pp-test="rapide"> (attentes de l'envoi divisées par 100).
//
// Un test qui échoue aujourd'hui exprime le comportement sûr attendu, porte un
// commentaire « // BUG : … » et l'étiquette @bug. Les autres forment la batterie de
// non-régression (ils passent aujourd'hui et doivent continuer à passer).
//
//   node node_modules/@playwright/test/cli.js test tests/audit-fenetres-politiques.spec.ts --workers=2 --reporter=line
//   … --grep @bug           seulement les constats
//   … --grep-invert @bug    seulement la non-régression
import { test as base, expect, chromium, type BrowserContext, type Page, type Worker } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
const EXT = join(ICI, "..", "extension");
const FAUX = join(ICI, "audit", "fenetres-politiques");
const PAGE = readFileSync(join(FAUX, "outlook-etablissement.html"), "utf8");
const JS = readFileSync(join(FAUX, "outlook-etablissement.js"), "utf8");
const CONNEXION = readFileSync(join(FAUX, "connexion.html"), "utf8");
const URL_OUTLOOK = "https://outlook.office.com/mail/";
const ID_ENVOYER = "splitButton-r200__primaryActionButton";

type Politique = Record<string, unknown>;
type Clic = { lib: string; id: string | null; zone: string; fenetre: string | null };
type Saisie = { type: string; cible: string; zone: string; fenetre: string | null; donnees: string | null };
type Mail = { a: string[]; cc: string[]; cci: string[]; objet: string; corps: string; html: string; importance: string };
type EtatFaux = {
  envoyes: Mail[]; clics: Clic[]; saisies: Saisie[]; fenetres: Array<{ id: string; ouverte: number; fermee?: number; titre: string }>;
  boiteEnvoi: Mail[]; brouillons: Mail[]; journal: Array<{ t: string; [k: string]: unknown }>; typed: string[]; debuts: number[]; clicsEnvoyerSousFenetre: number;
  keysHorsA: Array<{ key: string; cible: string; zone: string; fenetre: string | null }>; keysA: string[]; tentatives: number; reussis: number;
  aTexte: string; redactionOuverte: boolean; fenetresOuvertes: string[]; objetActuel: string | null; chargements: number; segEnCours: string;
};
type Resultat = { s: string; t?: number; d?: string; c?: number };
type Stock = { "pp-campagne"?: { curseur: number; total: number; resultats: Record<string, Resultat> }; "pp-deja"?: string[]; "pp-envois"?: number[]; "reprise-outlook"?: string; "base-outlook"?: string };

const test = base.extend<{ context: BrowserContext; sw: Worker; bloquees: string[] }>({
  // eslint-disable-next-line no-empty-pattern
  bloquees: async ({}, use) => { await use([]); },
  context: async ({ bloquees }, use) => {
    const dir = mkdtempSync(join(tmpdir(), "audit-fenetres-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      viewport: { width: 1280, height: 760 },
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    // Tout le reste est bloqué (et noté) ; les deux domaines servis le sont par ce banc.
    await context.route("**/*", (route) => { bloquees.push(route.request().url()); return route.abort(); });
    await context.route("https://outlook.office.com/**", (route) => {
      const p = new URL(route.request().url()).pathname;
      if (p === "/faux/outlook-etablissement.js") return route.fulfill({ status: 200, contentType: "text/javascript; charset=utf-8", body: JS });
      // ?lent=1 : vitesse réelle (sans le mode accéléré), pour les scénarios où l'ordre des événements compte.
      const lent = new URL(route.request().url()).searchParams.has("lent");
      return route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: lent ? PAGE.replace(' data-pp-test="rapide"', "") : PAGE });
    });
    await context.route("https://login.microsoftonline.com/**", (route) => route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: CONNEXION }));
    await use(context);
    await context.close();
    try { rmSync(dir, { recursive: true, force: true }); } catch { /* profil temporaire */ }
  },
  sw: async ({ context }, use) => {
    let [w] = context.serviceWorkers();
    if (!w) w = await context.waitForEvent("serviceworker");
    await use(w);
  },
});

test.describe.configure({ mode: "parallel" });

// ------------------------------------------------------------ outils
const pp = (page: Page) => page.locator("[data-pp-panneau]");
const ao = (page: Page) => page.locator('[data-testid="ao-panel"]');
const etat = (page: Page) => page.evaluate(() => (window as unknown as { __faux: { etat: () => EtatFaux } }).__faux.etat());
const stock = (sw: Worker) => sw.evaluate(() => (globalThis as any).chrome.storage.local.get(null)) as Promise<Stock>;
const statut = (s: Stock, i: number) => s["pp-campagne"]?.resultats?.[String(i)]?.s ?? "attente";
const clicsEnvoyer = (e: EtatFaux) => e.clics.filter((c) => c.id === ID_ENVOYER).length;
// Clics hors des commandes prévues : dans une fenêtre, une notification, un bandeau, la liste des messages, la
// recherche, une suggestion du champ « À », ou une commande marquée interdite (Cci-carnet, « Abandonner »…).
const clicsInterdits = (e: EtatFaux) => e.clics.filter((c) => !["redaction", "autre"].includes(c.zone));
// Saisies hors des champs de la rédaction (fenêtre, recherche, notification…).
const saisiesHors = (e: EtatFaux) => e.saisies.filter((s) => s.zone !== "redaction");

async function ouvrir(context: BrowserContext, sw: Worker, politique: Politique = {}, onglet: "envoi" | "adresses" = "envoi", vitesse = "fast", url = URL_OUTLOOK): Promise<Page> {
  await sw.evaluate(([o, v]) => (globalThis as any).chrome.storage.local.set({ "mega-onglet": o, reglages: JSON.stringify({ Speed: v }) }), [onglet, vitesse]);
  const page = context.pages()[0] ?? (await context.newPage());
  await page.addInitScript((p) => { (window as unknown as { __politique: Politique }).__politique = p; }, politique);
  await page.goto(url);
  await page.bringToFront();
  await pret(page);
  return page;
}
async function pret(page: Page) {
  await page.locator('html[data-mega-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached", timeout: 30000 });
  await page.locator('[data-testid="ao-panel"][data-ao-ready="1"]').waitFor({ state: "attached", timeout: 30000 });
}
async function politique(page: Page, o: Politique) {
  await page.evaluate((x) => (window as unknown as { __faux: { politique: (o: Politique) => void } }).__faux.politique(x), o);
}

type Envoi = { dest: string[]; objet?: string; texte?: string; pmin?: number; pmax?: number; max?: number; cci?: boolean; paquet?: number; importance?: boolean };
async function preparer(page: Page, o: Envoi) {
  const p = pp(page);
  await p.locator("#pp-dest").fill(o.dest.join("\n"));
  await p.locator("#pp-objet").fill(o.objet ?? "Invitation");
  await p.locator("#pp-texte").fill(o.texte ?? "Bonjour,\nVoici l'invitation à la journée portes ouvertes.");
  await p.locator("#pp-pmin").fill(String(o.pmin ?? 0));
  await p.locator("#pp-pmax").fill(String(o.pmax ?? 0));
  if (o.max != null) await p.locator("#pp-max").fill(String(o.max));
  if (o.cci) await p.locator("#pp-cci").check();
  if (o.paquet != null) await p.locator("#pp-paquet").fill(String(o.paquet));
  if (o.importance) await p.locator("#pp-importance").check();
}
// Note chaque message du panneau (Shadow DOM ouvert) : la fin d'un envoi ou d'un relevé se lit à son message.
async function observer(page: Page) {
  await page.evaluate(() => {
    const w = window as unknown as { __notices?: string[]; __noticesAO?: string[] };
    if (w.__notices) return;
    w.__notices = []; w.__noticesAO = [];
    const brancher = (hote: string, id: string, liste: string[]) => {
      const n = document.querySelector(hote)!.shadowRoot!.getElementById(id)!;
      new MutationObserver(() => liste.push(n.textContent || "")).observe(n, { childList: true, characterData: true, subtree: true });
    };
    brancher("[data-pp-panneau]", "pp-notice", w.__notices);
    brancher('[data-testid="ao-panel"]', "ao-notice", w.__noticesAO);
  });
}
const nbNotices = (page: Page, ao = false) => page.evaluate((x) => ((window as any)[x ? "__noticesAO" : "__notices"] as string[]).length, ao);
async function lancer(page: Page): Promise<number> {
  await observer(page);
  const avant = await nbNotices(page);
  await pp(page).locator("#pp-go").click();
  return avant;
}
async function arret(page: Page, avant: number, timeout = 60000): Promise<string> {
  await expect.poll(() => nbNotices(page), { timeout, message: "l'envoi ne s'arrête pas" }).toBeGreaterThan(avant);
  await expect(pp(page).locator("#pp-live")).toBeHidden();
  return (await pp(page).locator("#pp-notice").textContent()) ?? "";
}
async function envoyerTout(page: Page, o: Envoi, timeout = 60000): Promise<string> {
  await preparer(page, o);
  const avant = await lancer(page);
  return arret(page, avant, timeout);
}

// Relevé : l'utilisateur ouvre un nouveau message, clique dans « À », puis « Démarrer ».
async function demarrerReleve(page: Page) {
  await page.locator("#nouveau").click();
  await page.locator('[aria-label="À"][contenteditable="true"]').click();
  await observer(page);
  const avant = await nbNotices(page, true);
  await ao(page).locator("#ao-run").click();
  await expect(ao(page).locator("#ao-compact-view")).toBeVisible({ timeout: 15000 });
  return avant;
}
async function finReleve(page: Page, avant: number, timeout = 90000): Promise<string> {
  await expect.poll(() => nbNotices(page, true), { timeout, message: "le relevé ne s'arrête pas" }).toBeGreaterThan(avant);
  await expect(ao(page).locator("#ao-compact-view")).toBeHidden();
  return (await ao(page).locator("#ao-notice").textContent()) ?? "";
}


const D = (n: number, domaine = "exemple.fr", prefixe = "dest") => Array.from({ length: n }, (_, i) => `${prefixe}${i + 1}@${domaine}`);
async function attendreEnvoyes(page: Page, n: number, timeout = 30000) {
  await expect.poll(async () => (await etat(page)).envoyes.length, { timeout }).toBeGreaterThanOrEqual(n);
}
// Règle d'or : aucun clic dans une fenêtre / notification / bandeau, rien de tapé hors de la rédaction.
function regleDOr(e: EtatFaux) {
  expect(clicsInterdits(e), "clic hors des commandes prévues").toEqual([]);
  expect(saisiesHors(e), "saisie hors de la rédaction").toEqual([]);
}
// Le destinataire n'est pas « brûlé » : ni dans « Déjà envoyés », ni compté dans la limite sur 24 h.
async function pasBrule(sw: Worker, adresses: string[]) {
  const s = await stock(sw);
  for (const a of adresses) expect(s["pp-deja"] ?? [], `${a} mis dans « Déjà envoyés » alors que rien n'est parti`).not.toContain(a);
  expect((s["pp-envois"] ?? []).length, "compté dans la limite sur 24 h alors que rien n'est parti").toBe(0);
}
// Relevé : saisies ailleurs que dans le champ « À » (id 12), recherches commencées après l'ouverture d'une fenêtre.
const saisiesHorsA = (e: EtatFaux) => e.saisies.filter((x) => x.cible !== "12");
function recherchesApres(e: EtatFaux, idFenetre: string) {
  const f = e.fenetres.find((x) => x.id === idFenetre);
  expect(f, `la fenêtre « ${idFenetre} » s'est bien ouverte`).toBeTruthy();
  return e.debuts.filter((t) => t > f!.ouverte).length;
}

// ============================================================ contrôles du banc
test.describe("banc : la fausse page « établissement » se comporte comme Outlook", () => {
  test("envoi normal : un mail par destinataire, rien d'autre de cliqué ni tapé, aucune requête sortante", async ({ context, sw, bloquees }) => {
    const page = await ouvrir(context, sw);
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr", "carole@ecole.fr"] });
    expect(msg).toContain("Terminé : 3 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"], ["carole@ecole.fr"]]);
    for (const m of e.envoyes) { expect(m.cc).toEqual([]); expect(m.cci).toEqual([]); expect(m.objet).toBe("Invitation"); expect(m.corps).toContain("Voici l'invitation"); }
    expect(clicsEnvoyer(e)).toBe(3);
    regleDOr(e);
    expect(bloquees).toEqual([]);
  });

  test("relevé normal : les 8 contacts relevés, rien validé, « À » vide à la fin", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, {}, "adresses");
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    expect(msg).toContain("Terminé : 8 adresses");
    const e = await etat(page);
    expect(e.aTexte).toBe("");
    expect(e.journal.filter((j) => j.t === "pastille")).toEqual([]);
    expect(e.keysA.length).toBeGreaterThan(20); // le relevé a bien tapé dans « À »…
    expect(e.keysA.filter((k) => ["Enter", "Tab", ";", ","].includes(k))).toEqual([]); // … sans jamais valider
    expect(e.keysHorsA).toEqual([]);
    expect(saisiesHorsA(e)).toEqual([]);
    expect(clicsInterdits(e)).toEqual([]);
  });
});

// ============================================================ ENVOI — étiquette de sensibilité obligatoire
test.describe("envoi : étiquette de sensibilité obligatoire", () => {
  test("fenêtre « Étiquette requise » à l'envoi : aucun clic dedans, rien ne part, arrêt avec son texte", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { etiquette: "fenetre" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1); // un seul clic, jamais de second « Envoyer »
    regleDOr(e);
    expect(msg).toContain("Étiquette de sensibilité requise");
    expect(msg).not.toContain("Terminé");
    const s = await stock(sw);
    expect(statut(s, 0)).not.toBe("envoye");
    expect(statut(s, 1)).toBe("attente"); // le suivant n'est pas tenté
    await expect(pp(page).locator("#pp-go")).toHaveText("Reprendre l'envoi");
  });

  test("fenêtre « Étiquette requise » : le destinataire n'est pas classé « Déjà envoyés » ni compté sur 24 h", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : Outlook refuse explicitement (rédaction ouverte, fenêtre d'étiquette) mais l'extension classe le
    // destinataire « à vérifier », l'ajoute à « Déjà envoyés » (il ne recevra jamais le mail) et le compte dans la
    // limite du jour (outlook.js:727-731 → moteur.js:316 et 399).
    const page = await ouvrir(context, sw, { etiquette: "fenetre" });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    expect((await etat(page)).envoyes).toEqual([]);
    await pasBrule(sw, ["alice@exemple.fr"]);
  });

  test("bouton « Envoyer » désactivé tant qu'aucune étiquette n'est choisie : arrêt sans brûler le destinataire, raison dite", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : l'extension clique un « Envoyer » aria-disabled (qui ne peut rien envoyer), attend 60 s, classe « à
    // vérifier », ajoute à « Déjà envoyés » et compte sur 24 h ; le message ne dit pas pourquoi (outlook.js:723-731).
    const page = await ouvrir(context, sw, { etiquette: "desactive" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    regleDOr(e); // jamais « Réglez votre sensibilité » ni un choix d'étiquette
    expect(msg).toMatch(/désactivé|étiquette/i);
    expect(statut(await stock(sw), 0)).toBe("attente");
    await pasBrule(sw, ["alice@exemple.fr"]);
  });

  test("étiquette exigée par un bandeau dans la rédaction : le message d'arrêt dit pourquoi", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : seul le texte d'une fenêtre est repris ; un refus affiché en bandeau dans la rédaction donne
    // « la rédaction est restée ouverte. » sans la raison (outlook.js:729-731).
    const page = await ouvrir(context, sw, { etiquette: "barre" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    expect((await etat(page)).envoyes).toEqual([]);
    regleDOr(await etat(page));
    expect(msg).toContain("étiquette de sensibilité");
  });

  test("fenêtre de choix d'étiquette ouverte dès la rédaction et restée : arrêt, rien tapé ni cliqué dedans", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { etiquette: "ouverture" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(0);
    regleDOr(e);
    expect(msg).toMatch(/Outlook affiche une fenêtre par-dessus.*étiquette/);
    expect(statut(await stock(sw), 0)).toBe("attente");
    await pasBrule(sw, ["alice@exemple.fr"]);
  });

  test("Cci : fenêtre « Étiquette requise » pour le paquet : aucun clic dedans, rien ne part, le paquet suivant n'est pas tenté", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { etiquette: "fenetre" });
    await envoyerTout(page, { dest: D(8), cci: true, paquet: 4, max: 1000 });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
    const s = await stock(sw);
    expect([4, 5, 6, 7].map((i) => statut(s, i))).toEqual(["attente", "attente", "attente", "attente"]);
  });
});

// ============================================================ ENVOI — destinataires externes
test.describe("envoi : avertissements « destinataires externes »", () => {
  test("MailTip dans la rédaction (non bloquant) : le mail part, aucun bouton du MailTip cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "mailtip" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"]]);
    regleDOr(e);
  });

  test("fenêtre de confirmation à l'envoi (« Ne pas envoyer » / « Envoyer quand même ») : jamais cliquée, rien ne part", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "fenetre" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
    expect(msg).toContain("Contrôle des destinataires externes");
    const s = await stock(sw);
    expect(statut(s, 0)).not.toBe("envoye");
    expect(statut(s, 1)).toBe("attente");
  });

  test("même fenêtre sans aria-modal (non modale) : jamais cliquée, rien ne part", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "fenetre-non-modale" });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    regleDOr(e);
  });

  test("interne seulement : pas de fenêtre, le mail part ; puis un externe : arrêt sur la fenêtre", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "fenetre" });
    await envoyerTout(page, { dest: ["carole@ecole.fr", "alice@exemple.fr", "denis@ecole.fr"] });
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a)).toEqual([["carole@ecole.fr"]]);
    expect(clicsEnvoyer(e)).toBe(2);
    regleDOr(e);
    expect(statut(await stock(sw), 2)).toBe("attente");
  });

  test("l'utilisateur clique lui-même « Envoyer quand même » pendant l'attente : compté une fois, jamais renvoyé", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "fenetre" });
    await preparer(page, { dest: ["alice@exemple.fr", "bob@ecole.fr"] });
    // L'utilisateur, devant la fenêtre, clique « Envoyer quand même » 100 ms après son apparition.
    await page.evaluate(() => new MutationObserver((_, obs) => {
      const b = [...document.querySelectorAll('[data-fenetre="smart-alert"] button')].find((x) => x.textContent === "Envoyer quand même") as HTMLButtonElement | undefined;
      if (b) { obs.disconnect(); setTimeout(() => b.click(), 100); }
    }).observe(document.body, { childList: true, subtree: true }));
    const avant = await lancer(page);
    const msg = await arret(page, avant);
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["alice@exemple.fr", "bob@ecole.fr"]);
    expect(clicsInterdits(e).length).toBe(1); // le seul clic dans la fenêtre est celui de l'utilisateur
    expect(clicsEnvoyer(e)).toBe(2);
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
  });

  test("« Envoyer quand même » cliqué par l'utilisateur APRÈS l'arrêt « à vérifier » : la reprise ne le renvoie pas", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "fenetre" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@ecole.fr"] });
    expect(msg).toContain("ne sera pas renvoyé");
    await page.locator('[data-fenetre="smart-alert"] button', { hasText: "Envoyer quand même" }).click();
    await attendreEnvoyes(page, 1);
    const avant = await lancer(page);
    await arret(page, avant);
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["alice@exemple.fr", "bob@ecole.fr"]); // alice une seule fois
    const s = await stock(sw);
    expect([statut(s, 0), statut(s, 1)]).toEqual(["verifier", "envoye"]);
  });

  test("Cci : fenêtre de confirmation pour un paquet externe : jamais cliquée, rien ne part", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { externe: "fenetre" });
    await envoyerTout(page, { dest: D(6), cci: true, paquet: 3, max: 1000 });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
  });
});

// ============================================================ ENVOI — DLP
test.describe("envoi : conseils de stratégie DLP", () => {
  test("DLP bloquant à l'envoi (« OK ») : jamais cliqué, rien ne part, arrêt", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { dlp: "bloquant" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
    expect(msg).toContain("Message bloqué");
  });

  test("DLP « Ignorer avec justification » : rien tapé dans la justification, « Ignorer et envoyer » jamais cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { dlp: "justification" });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    regleDOr(e);
    expect(await page.locator("#justification").inputValue()).toBe("");
  });

  test("conseil de stratégie non bloquant dans la rédaction : le mail part, « Ignorer » / « Signaler » jamais cliqués", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { dlp: "conseil" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    regleDOr(await etat(page));
  });

  test("DLP bloquant affiché en bandeau rouge à l'envoi : rien ne part, « Ignorer » du bandeau jamais cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { dlp: "barre-bloquante" });
    await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
  });
});

// ============================================================ ENVOI — avertissements d'Outlook
test.describe("envoi : avertissements d'Outlook (objet, pièce jointe, noms, nombre de destinataires)", () => {
  test("rappel de pièce jointe oubliée : son bouton « Envoyer » (même libellé) n'est jamais cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { pieceJointe: true });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"], texte: "Bonjour,\nVeuillez trouver ci-joint le programme." });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
    expect(msg).toContain("pièce jointe");
  });

  test("rappel de pièce jointe affiché DANS la rédaction (avant le vrai « Envoyer » dans le DOM) : jamais cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { pieceJointe: "dans-redaction" });
    await envoyerTout(page, { dest: ["alice@exemple.fr"], texte: "Voir la pièce jointe." });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
  });

  test("« Ce message n'a pas d'objet » (objet vidé par un complément) : le « Envoyer » de la fenêtre n'est jamais cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "apres-corps", duree: 200, effet: "vide-objet" } });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    regleDOr(e);
  });

  test("adresse que Outlook ne résout pas (pastille soulignée) : jamais de clic dans « Vérifiez les noms », rien ne part", async ({ context, sw }) => {
    const page = await ouvrir(context, sw);
    await envoyerTout(page, { dest: ["jean..dupont@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e); // ni « OK », ni « Vérifier les noms » du ruban
  });

  test("adresse non résolue : repérée avant « Envoyer » (pastille non valide), le destinataire n'est pas brûlé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : remplirA accepte n'importe quelle pastille (outlook.js:367-372), même marquée non valide ; Outlook refuse
    // ensuite à l'envoi et l'adresse finit « à vérifier » + « Déjà envoyés » + comptée sur 24 h.
    const page = await ouvrir(context, sw);
    const msg = await envoyerTout(page, { dest: ["jean..dupont@exemple.fr"] });
    expect(clicsEnvoyer(await etat(page))).toBe(0);
    expect(msg).toMatch(/valide|résolu|Vérifiez les noms/i);
    await pasBrule(sw, ["jean..dupont@exemple.fr"]);
  });

  test("trop de destinataires pour l'établissement (Cci 30, limite 20) : jamais de clic dans la fenêtre, rien ne part", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { maxDestinataires: 20 });
    await envoyerTout(page, { dest: D(40), cci: true, paquet: 30, max: 1000 });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
  });

  test("trop de destinataires : MailTip vu avant « Envoyer » → arrêt sans brûler les 30 destinataires du paquet", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : le MailTip « trop de destinataires » est ignoré, « Envoyer » est cliqué, Outlook refuse, et les 30
    // adresses du paquet passent « à vérifier » + « Déjà envoyés » + 30 comptées sur 24 h (moteur.js:316, 389-399).
    const page = await ouvrir(context, sw, { maxDestinataires: 20 });
    await envoyerTout(page, { dest: D(30), cci: true, paquet: 30, max: 1000 });
    expect((await etat(page)).envoyes).toEqual([]);
    await pasBrule(sw, D(30));
  });
});

// ============================================================ ENVOI — « Annuler l'envoi »
test.describe("envoi : délai « Annuler l'envoi »", () => {
  test("délai de 1 s avec compte à rebours : chacun reçoit un seul mail, « Annuler » jamais cliqué, même sans pause", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { annulerEnvoi: 1000 });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr", "carole@exemple.fr"] });
    expect(msg).toContain("Terminé : 3 mail(s) envoyé(s).");
    await attendreEnvoyes(page, 3, 10000);
    await page.waitForTimeout(1500);
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["alice@exemple.fr", "bob@exemple.fr", "carole@exemple.fr"]);
    expect(e.journal.filter((j) => j.t === "envoi-annule")).toEqual([]);
    regleDOr(e);
  });

  test("page rechargée pendant le délai d'annulation (mise à jour d'Outlook) : le mail perdu n'est pas compté envoyé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : la fermeture de la rédaction vaut « envoyé » (outlook.js:727-728) alors que le bandeau « Envoi en
    // cours… Annuler » court encore ; rechargée pendant ce délai, la page ne fait jamais partir le mail (variante
    // plausible), qui reste « envoyé » et dans « Déjà envoyés » : il ne sera jamais renvoyé.
    const page = await ouvrir(context, sw, { annulerEnvoi: 1500, maj: { rechargerApresEnvoi: 300, une: true } });
    await preparer(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"], pmin: 100, pmax: 100 });
    await lancer(page);
    await expect.poll(async () => (await page.evaluate(() => (window as any).__faux?.etat().chargements ?? 0).catch(() => 0)), { timeout: 20000 }).toBeGreaterThanOrEqual(2);
    await pret(page);
    await page.waitForTimeout(2500);
    const e = await etat(page);
    const s = await stock(sw);
    expect(e.envoyes.map((m) => m.a[0])).not.toContain("alice@exemple.fr"); // le mail n'est jamais parti (contrôle du scénario)
    expect(statut(s, 0)).not.toBe("envoye");
  });
});

// ============================================================ ENVOI — compléments (Letsignit et autres)
test.describe("envoi : compléments (Letsignit…)", () => {
  test("fenêtre de complément non bloquante (sans aria-modal) : le mail part, rien cliqué dedans", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "ouverture", duree: null, modale: false } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    regleDOr(await etat(page));
  });

  test("Letsignit passager à l'ouverture (300 ms) : on attend, puis un seul mail, contenu exact", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "ouverture", duree: 300 } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"]]);
    expect(e.clicsEnvoyerSousFenetre).toBe(0);
    regleDOr(e);
  });

  test("Letsignit passager à l'ouverture AVEC piège à focus (300 ms) : on attend, puis un seul mail, contenu exact", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "ouverture", duree: 300, piege: "button" } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => [m.a, m.objet])).toEqual([[["alice@exemple.fr"], "Invitation"], [["bob@exemple.fr"], "Invitation"]]);
    for (const m of e.envoyes) expect(m.corps).toContain("Voici l'invitation");
    regleDOr(e);
  });

  test("Letsignit bloquant à l'ouverture, qui reste : arrêt avec son nom, rien tapé dans la rédaction ni cliqué dans la fenêtre", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "ouverture", duree: null, piege: "button" } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(0);
    regleDOr(e);
    expect(e.saisies.filter((x) => x.type === "paste")).toEqual([]); // rien collé sous la fenêtre
    expect(msg).toContain("Letsignit");
    expect(statut(await stock(sw), 0)).toBe("attente");
  });

  test("Letsignit tardif (après « À »), qui reste : arrêt, rien ne part, rien cliqué dans la fenêtre", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "apres-a", duree: null } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(0);
    regleDOr(e);
    expect(msg).toContain("Letsignit");
  });

  test("Letsignit qui se rouvre sans cesse pendant la rédaction : chaque destinataire reçoit un seul mail, exact", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "clignote", duree: 40, periode: 90 } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr", "carole@exemple.fr"] });
    expect(msg).toContain("Terminé : 3 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"], ["carole@exemple.fr"]]);
    for (const m of e.envoyes) { expect(m.objet).toBe("Invitation"); expect(m.corps).toContain("Voici l'invitation"); expect(m.cc).toEqual([]); }
    regleDOr(e);
  });

  test("Letsignit à l'envoi (vérification de 300 ms) : le mail part une fois, compté une fois", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "envoi", duree: 300 } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["alice@exemple.fr", "bob@exemple.fr"]);
    regleDOr(e);
  });

  test("Letsignit à l'envoi qui ne rend jamais la main : « à vérifier », jamais renvoyé ni cliqué, arrêt", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "envoi", duree: null } });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
    expect(msg).toContain("ne sera pas renvoyé");
    const s = await stock(sw);
    expect([statut(s, 0), statut(s, 1)]).toEqual(["verifier", "attente"]);
    // L'utilisateur ferme la fenêtre et abandonne le brouillon, puis reprend : le destinataire incertain n'est jamais renvoyé.
    await politique(page, { letsignit: null });
    await page.evaluate(() => { (window as any).__faux.fermerFenetre("letsignit-envoi"); document.getElementById("abandonner")!.click(); });
    const avant = await lancer(page);
    await arret(page, avant);
    const e2 = await etat(page);
    expect(e2.envoyes.map((m) => m.a[0])).toEqual(["bob@exemple.fr"]);
  });

  test("« Retirer le texte de la signature » + Letsignit qui remet sa signature en se refermant : le mail ne part pas avec ce texte", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : aucun contrôle juste avant le clic : envoyer() attend que la fenêtre Letsignit se ferme
    // (outlook.js:720-721) puis clique sans revérifier ; le dernier retrait de signature (moteur.js:305-309) a eu
    // lieu AVANT. Le mail part avec « Pierre TESTEUR / École de test ».
    const page = await ouvrir(context, sw, { letsignit: { quand: "apres-corps", duree: 200, effet: "reinsere-signature" } });
    await pp(page).locator("#pp-sans-signature").check();
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    for (const m of e.envoyes) expect(m.corps).not.toContain("Pierre TESTEUR");
  });

  test("complément tardif qui ajoute un destinataire en Cc : le mail ne part pas avec un destinataire en trop", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : « À » est contrôlé juste après le collage (outlook.js:405-406), jamais Cc/Cci, et rien n'est revérifié
    // avant « Envoyer » : le mail part aussi à archives@ecole.fr (E1).
    const page = await ouvrir(context, sw, { letsignit: { quand: "apres-corps", duree: 200, effet: "ajoute-cc" } });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    for (const m of e.envoyes) expect([...m.a, ...m.cc, ...m.cci]).toEqual(["alice@exemple.fr"]);
  });

  test("complément tardif qui réécrit le corps : le mail ne part pas sans le message", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : le corps n'est vérifié qu'au collage (outlook.js:555-565) ; réécrit ensuite par le complément, il part
    // sans le texte prévu (E3).
    const page = await ouvrir(context, sw, { letsignit: { quand: "apres-corps", duree: 200, effet: "reecrit-corps" } });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    for (const m of e.envoyes) expect(m.corps).toContain("Voici l'invitation");
  });

  test("complément tardif qui vide l'objet : « Envoyer » n'est pas cliqué (l'objet n'est plus celui prévu)", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : pas de contrôle final de l'objet : « Envoyer » est cliqué avec un objet vide ; seule la fenêtre
    // « Ce message n'a pas d'objet » d'Outlook empêche le départ, et le destinataire est brûlé « à vérifier ».
    const page = await ouvrir(context, sw, { letsignit: { quand: "apres-corps", duree: 200, effet: "vide-objet" } });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    expect(clicsEnvoyer(await etat(page))).toBe(0);
    await pasBrule(sw, ["alice@exemple.fr"]);
  });

  test("fenêtre de complément à zone de saisie (piège à focus) qui s'ouvre au clic dans l'Objet : rien n'y est tapé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : coller() → insererTexte() fait execCommand("insertText") sur l'élément qui a le focus sans vérifier que
    // c'est l'Objet (outlook.js:209-221, 339-353) : le piège à focus de la fenêtre y ramène le focus et l'objet
    // « Invitation » est tapé DANS la fenêtre ; l'arrêt dit ensuite « L'objet collé ne correspond pas ».
    const page = await ouvrir(context, sw, { letsignit: { quand: "objet", duree: null, piege: "#letsignit-champ", champ: "Décrivez ce que vous voulez écrire" } });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(saisiesHors(e)).toEqual([]);
    expect(await page.locator("#letsignit-champ").inputValue()).toBe("");
  });
});

// ============================================================ ENVOI — échecs signalés, limitation, boîte pleine
test.describe("envoi : échec d'envoi, limitation, boîte pleine", () => {
  test("« Le message n'a pas pu être envoyé » (notification après fermeture, brouillon gardé) : pas compté envoyé, arrêt", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : rédaction fermée = « envoyé » (outlook.js:727-728, moteur.js:314) ; la notification d'échec qui suit n'est
    // jamais lue : le mail est compté envoyé, mis dans « Déjà envoyés », et la campagne continue (« Terminé : 3 »).
    const page = await ouvrir(context, sw, { echec: "toast", echecJusqua: 0 });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr", "carole@exemple.fr"], pmin: 100, pmax: 100 });
    const e = await etat(page);
    expect(e.brouillons.map((m) => m.a[0])).toEqual(["alice@exemple.fr"]); // contrôle du scénario : il n'est pas parti
    expect(statut(await stock(sw), 0)).not.toBe("envoye");
    expect(clicsEnvoyer(e)).toBe(1); // arrêt avant le mail suivant
    expect(msg).not.toContain("Terminé");
  });

  test("limitation « trop de messages » après le 1er mail : les mails refusés ne sont pas comptés envoyés, arrêt au 1er refus", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause : chaque refus (notification « trop de messages ») est compté envoyé ; la campagne finit
    // « Terminé : 4 mail(s) envoyé(s) » alors qu'un seul est parti, et les 3 autres ne seront jamais renvoyés.
    const page = await ouvrir(context, sw, { limitation: 1 });
    const msg = await envoyerTout(page, { dest: D(4), pmin: 100, pmax: 100 });
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["dest1@exemple.fr"]);
    const s = await stock(sw);
    expect([1, 2, 3].map((i) => statut(s, i))).not.toContain("envoye");
    expect(clicsEnvoyer(e)).toBe(2);
    expect(msg).not.toContain("Terminé");
  });

  test("brouillon rouvert avec « Ce message n'a pas pu être envoyé » : pas compté envoyé, message d'arrêt juste", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause ; en plus l'arrêt au mail suivant dit « Un message est déjà ouvert en rédaction » (trompeur).
    const page = await ouvrir(context, sw, { echec: "reouvre", echecJusqua: 0 });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"], pmin: 100, pmax: 100 });
    expect((await etat(page)).envoyes).toEqual([]);
    expect(statut(await stock(sw), 0)).not.toBe("envoye");
    expect(msg).toMatch(/pas pu être envoyé|échec/i);
  });

  test("Cci : échec signalé pour un paquet : ses destinataires ne sont pas comptés servis", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause, à l'échelle d'un paquet : 5 destinataires comptés servis et « Déjà envoyés » sans rien reçu.
    const page = await ouvrir(context, sw, { echec: "toast", echecJusqua: 0 });
    await envoyerTout(page, { dest: D(10), cci: true, paquet: 5, max: 1000, pmin: 100, pmax: 100 });
    const s = await stock(sw);
    expect([0, 1, 2, 3, 4].map((i) => statut(s, i))).not.toContain("envoye");
    for (const a of D(5)) expect(s["pp-deja"] ?? []).not.toContain(a);
  });

  test("boîte pleine (bandeau + refus à l'envoi) : rien ne part, aucun bouton du bandeau cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { boitePleine: true });
    await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
  });

  test("boîte pleine : refus explicite d'Outlook → le destinataire n'est pas brûlé, la raison est dite", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : refus explicite (bandeau « boîte aux lettres pleine ») classé « à vérifier » : « Déjà envoyés » + limite
    // du jour, et le message ne dit pas la raison (outlook.js:727-731).
    const page = await ouvrir(context, sw, { boitePleine: true });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    expect(msg).toMatch(/pleine/);
    await pasBrule(sw, ["alice@exemple.fr"]);
  });
});

// ============================================================ ENVOI — session expirée
test.describe("envoi : session expirée", () => {
  test("redirection vers login.microsoftonline.com juste après « Envoyer » : au retour « à vérifier », jamais renvoyé, rien fait sur la page de connexion", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { session: "redirection-envoi" });
    await preparer(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    await lancer(page);
    await page.waitForURL(/login\.microsoftonline\.com/, { timeout: 20000 });
    await page.waitForTimeout(800);
    const surConnexion = await page.evaluate(() => ({ panneaux: document.querySelectorAll("[data-pp-panneau],[data-testid='ao-panel']").length, saisies: (window as any).__saisies as string[] }));
    expect(surConnexion).toEqual({ panneaux: 0, saisies: [] });
    await page.addInitScript(() => { try { sessionStorage.setItem("__politique", JSON.stringify({ session: null })); } catch { /* */ } });
    await page.goto(URL_OUTLOOK);
    await pret(page);
    let s = await stock(sw);
    expect([statut(s, 0), statut(s, 1)]).toEqual(["verifier", "attente"]);
    expect((await etat(page)).envoyes).toEqual([]);
    await expect(pp(page).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await page.waitForTimeout(1500);
    expect((await etat(page)).tentatives).toBe(1); // rien ne repart tout seul après la reconnexion
    const avant = await lancer(page);
    await arret(page, avant);
    s = await stock(sw);
    expect((await etat(page)).envoyes.map((m) => m.a[0])).toEqual(["bob@exemple.fr"]);
    expect(statut(s, 0)).toBe("verifier");
  });

  test("redirection pendant le remplissage (avant « Envoyer ») : au retour le destinataire est en attente, puis reçoit UN mail", async ({ context, sw }) => {
    // À vitesse réelle : la redirection survient pendant le collage de l'objet, bien avant « Envoyer ».
    const page = await ouvrir(context, sw, { session: "redirection-objet" }, "envoi", "fast", URL_OUTLOOK + "?lent=1");
    await preparer(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    await lancer(page);
    await page.waitForURL(/login\.microsoftonline\.com/, { timeout: 20000 });
    await page.goto(URL_OUTLOOK);
    await pret(page);
    expect(statut(await stock(sw), 0)).toBe("attente");
    expect((await etat(page)).tentatives).toBe(0);
    const avant = await lancer(page);
    const msg = await arret(page, avant);
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    expect((await etat(page)).envoyes.map((m) => m.a[0])).toEqual(["alice@exemple.fr", "bob@exemple.fr"]);
  });

  test("fenêtre « Votre session a expiré » dans la page : arrêt, « Se connecter » jamais cliqué, rien ne part", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { session: "fenetre-ouverture" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(0);
    regleDOr(e);
    expect(msg).toContain("session a expiré");
    expect(page.url()).toContain("outlook.office.com");
  });

  test("écran « session expirée » qui remplace Outlook juste après « Envoyer » : le mail n'est pas compté envoyé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : la rédaction a disparu (l'application est remplacée par l'écran de session) et c'est compté « envoyé »
    // (outlook.js:727-728) ; l'arrêt suivant dit « Bouton « Nouveau message » introuvable ».
    const page = await ouvrir(context, sw, { session: "ecran-envoi" });
    await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect((await etat(page)).envoyes).toEqual([]);
    expect(statut(await stock(sw), 0)).not.toBe("envoye");
  });
});

// ============================================================ ENVOI — hors ligne
test.describe("envoi : hors ligne", () => {
  test("hors ligne dès le départ (Outlook garde les mails en Boîte d'envoi) : arrêt sans cliquer « Envoyer », raison dite", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : aucun contrôle de la connexion (navigator.onLine, bandeau « hors connexion ») dans moteur.js / outlook.js :
    // l'extension ouvre, remplit, clique « Envoyer » ; les 2 mails partent en Boîte d'envoi et sont comptés envoyés.
    const page = await ouvrir(context, sw, { horsLigne: "boite-envoi" });
    await preparer(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    await context.setOffline(true);
    const avant = await lancer(page);
    const msg = await arret(page, avant);
    const e = await etat(page);
    await context.setOffline(false);
    expect(clicsEnvoyer(e)).toBe(0);
    expect(msg).toMatch(/connexion|hors ligne/i);
    const s = await stock(sw);
    expect([statut(s, 0), statut(s, 1)]).toEqual(["attente", "attente"]);
  });

  test("connexion perdue pendant la pause entre deux mails : le mail suivant n'est pas tenté", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause : la campagne continue hors ligne et compte envoyés des mails restés en Boîte d'envoi.
    const page = await ouvrir(context, sw, { horsLigne: "boite-envoi" });
    await preparer(page, { dest: D(3), pmin: 150, pmax: 150 });
    const avant = await lancer(page);
    await attendreEnvoyes(page, 1);
    await context.setOffline(true);
    const msg = await arret(page, avant);
    const e = await etat(page);
    await context.setOffline(false);
    expect(clicsEnvoyer(e)).toBe(1);
    expect(msg).toMatch(/connexion|hors ligne/i);
  });

  test("hors ligne, Outlook refuse d'envoyer : arrêt avant « Envoyer », le destinataire n'est pas brûlé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause ; ici le refus laisse la rédaction ouverte : « à vérifier » + « Déjà envoyés » + limite du jour.
    const page = await ouvrir(context, sw, { horsLigne: "refus" });
    await preparer(page, { dest: ["alice@exemple.fr"] });
    await context.setOffline(true);
    const avant = await lancer(page);
    await arret(page, avant);
    const e = await etat(page);
    await context.setOffline(false);
    expect(e.envoyes).toEqual([]);
    expect(clicsEnvoyer(e)).toBe(0);
    await pasBrule(sw, ["alice@exemple.fr"]);
  });
});

// ============================================================ ENVOI — mise à jour d'Outlook, notifications, rappels, diagnostic
test.describe("envoi : mise à jour d'Outlook, notifications, rappels, diagnostic", () => {
  test("bandeau « nouvelle version disponible » : les mails partent, « Actualiser » / « Envoyer des commentaires » jamais cliqués", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { maj: "bandeau" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(clicsEnvoyer(e)).toBe(2);
    regleDOr(e);
  });

  test("fenêtre « Outlook a été mis à jour » à l'ouverture d'un message : arrêt, rien ne part, rien cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { maj: "fenetre" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    regleDOr(e);
    expect(msg).toContain("Outlook a été mis à jour");
    expect(e.chargements).toBe(1); // pas d'« Actualiser » cliqué
  });

  test("page rechargée par une mise à jour pendant la campagne : rien ne repart tout seul, reprise sans doublon", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { maj: { rechargerApresEnvoi: 300, une: true } });
    await preparer(page, { dest: D(3), pmin: 100, pmax: 100 });
    await lancer(page);
    await expect.poll(async () => (await page.evaluate(() => (window as any).__faux?.etat().chargements ?? 0).catch(() => 0)), { timeout: 20000 }).toBeGreaterThanOrEqual(2);
    await pret(page);
    await page.waitForTimeout(1500);
    let e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["dest1@exemple.fr"]);
    expect(e.tentatives).toBe(1); // rien ne repart tout seul
    const avant = await lancer(page);
    await arret(page, avant);
    e = await etat(page);
    expect(e.envoyes.map((m) => m.a[0])).toEqual(["dest1@exemple.fr", "dest2@exemple.fr", "dest3@exemple.fr"]);
  });

  test("notification « Nouveau message de … » pendant la rédaction : jamais cliquée, chaque mail part une fois", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { toastNouveauMessage: "apres-a" });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"]]);
    regleDOr(e);
  });

  test("rappel de réunion affiché (fenêtre non modale « Ignorer / Répéter ») : les mails partent, le rappel n'est pas touché", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { rappel: true });
    const msg = await envoyerTout(page, { dest: ["alice@exemple.fr", "bob@exemple.fr"] });
    expect(msg).toContain("Terminé : 2 mail(s) envoyé(s).");
    const e = await etat(page);
    regleDOr(e);
    expect(e.fenetresOuvertes).toContain("rappel");
  });

  test("fenêtre modale qui s'ouvre pendant la pause et reste : le mail suivant n'est pas tenté", async ({ context, sw }) => {
    const page = await ouvrir(context, sw);
    await preparer(page, { dest: D(3), pmin: 150, pmax: 150 });
    const avant = await lancer(page);
    await attendreEnvoyes(page, 1);
    await page.evaluate(() => (window as any).__faux.ouvrirFenetre({ id: "nouveautes", titre: "Découvrez Copilot dans Outlook", texte: "Rédigez plus vite.", boutons: [{ lib: "Essayer" }, { lib: "Plus tard" }] }));
    const msg = await arret(page, avant);
    const e = await etat(page);
    expect(e.envoyes.length).toBe(1);
    expect(clicsEnvoyer(e)).toBe(1);
    regleDOr(e);
    expect(msg).toContain("Découvrez Copilot");
  });

  test("diagnostic pendant qu'une fenêtre Letsignit couvre le brouillon : rien n'est cliqué dans la fenêtre", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "ouverture", duree: null, piege: "button" } });
    await pp(page).locator("#pp-gear").click();
    await pp(page).locator("#pp-diag").click();
    await expect(pp(page).locator("#pp-notice2")).toContainText(/Diagnostic téléchargé|Diagnostic :/, { timeout: 30000 });
    const e = await etat(page);
    expect(clicsInterdits(e)).toEqual([]);
    expect(saisiesHors(e)).toEqual([]);
    expect(e.envoyes).toEqual([]);
  });

  test("diagnostic avec un rappel de réunion affiché : le bouton « Ignorer » du rappel n'est jamais cliqué", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : le diagnostic cherche « Ignorer » dans TOUTE la page (SEL.ignorer, outlook.js:46 et 795-800) alors que la
    // rédaction réelle a « Abandonner » : il clique « Ignorer » du rappel (rappel supprimé) — et à défaut tout bouton
    // « OK » / « Ignorer » de la première fenêtre venue.
    const page = await ouvrir(context, sw, { rappel: true });
    await pp(page).locator("#pp-gear").click();
    await pp(page).locator("#pp-diag").click();
    await expect(pp(page).locator("#pp-notice2")).toContainText(/Diagnostic téléchargé|Diagnostic :/, { timeout: 30000 });
    const e = await etat(page);
    expect(clicsInterdits(e)).toEqual([]);
    expect(e.fenetresOuvertes).toContain("rappel");
  });
});

// ============================================================ ENVOI — fenêtres qui contiennent des champs « À » / « Objet »
test.describe("envoi : fenêtre qui contient elle-même des champs de destinataires", () => {
  test("« Ajouter des destinataires » (carnet d'adresses) ouvert par-dessus la rédaction : rien n'est envoyé sous la fenêtre", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : dialogueBloquant() écarte toute fenêtre qui contient un champ « À » ou « Objet » (outlook.js:117-118),
    // pensant que c'est la rédaction « isolée » : l'extension remplit et ENVOIE sous le carnet d'adresses, puis prend
    // son champ « À » pour la rédaction restée ouverte (« à vérifier » alors que le mail est parti).
    const page = await ouvrir(context, sw, { carnetSurRedaction: true });
    await envoyerTout(page, { dest: ["alice@exemple.fr"] });
    const e = await etat(page);
    expect(saisiesHors(e)).toEqual([]); // rien tapé dans les champs du carnet
    expect(e.envoyes).toEqual([]);
    expect(e.clicsEnvoyerSousFenetre).toBe(0);
  });
});

// ============================================================ RELEVÉ — fenêtres pendant le relevé
test.describe("relevé : fenêtres qui s'ouvrent pendant le relevé", () => {
  test("fenêtre Letsignit (piège à focus, comme ms-Modal) : plus aucune recherche tapée, « À » vidé, arrêt qui le dit", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : adapter-dom.js ne détecte aucune fenêtre ; taper() insère par la sélection même quand le focus est
    // dans la fenêtre (adapter-dom.js:73-97) : une recherche de plus est tapée SOUS la fenêtre, puis focus() échoue
    // et le moteur s'arrête sans effacer (core/moteur.js:542 n'efface que si le focus revient) : « d » reste dans « À ».
    const page = await ouvrir(context, sw, { releve: { type: "letsignit", apres: 3, piege: "button" } }, "adresses");
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant);
    const e = await etat(page);
    expect(saisiesHorsA(e)).toEqual([]);
    expect(clicsInterdits(e)).toEqual([]);
    expect(recherchesApres(e, "letsignit"), "recherches tapées sous la fenêtre").toBe(0);
    expect(e.aTexte, "texte laissé dans « À »").toBe("");
    expect(msg).toMatch(/fenêtre|Letsignit/);
  });

  test("fenêtre Letsignit passagère (1,5 s) : « À » ne garde jamais le texte tapé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause : arrêt « le curseur a quitté le champ « À » » et la recherche reste tapée dans « À »
    // après la fermeture de la fenêtre (R2).
    const page = await ouvrir(context, sw, { releve: { type: "letsignit", apres: 3, piege: "button", duree: 1500 } }, "adresses");
    const avant = await demarrerReleve(page);
    await finReleve(page, avant);
    await page.waitForTimeout(2000);
    const e = await etat(page);
    expect(e.fenetresOuvertes).toEqual([]);
    expect(e.aTexte, "texte laissé dans « À »").toBe("");
  });

  test("« Pause » demandée pendant qu'une fenêtre Letsignit (piège à focus) est affichée : « À » est vidé quand même", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause : à la pause, focus() échoue (le piège garde le focus) et core/moteur.js:542 n'efface pas.
    // La fenêtre s'ouvre juste après la 1re lettre d'une recherche et la pause est demandée au même instant.
    const page = await ouvrir(context, sw, { releve: { type: "letsignit", apres: 2, moment: "debut-recherche", piege: "button", pause: true } }, "adresses");
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant);
    expect(msg).toContain("En pause");
    const e = await etat(page);
    expect(saisiesHorsA(e)).toEqual([]);
    expect(e.aTexte, "texte laissé dans « À »").toBe("");
  });

  test("fenêtre avec zone de saisie et piège à focus : aucune lettre n'est tapée dans la fenêtre", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : insertText() fait execCommand("insertText") sur l'élément qui a le focus sans vérifier que c'est le champ
    // « À » (adapter-dom.js:73-87) : la lettre suivante est tapée DANS la fenêtre (R4).
    const page = await ouvrir(context, sw, { releve: { type: "champ", apres: 3 } }, "adresses");
    const avant = await demarrerReleve(page);
    await finReleve(page, avant);
    const e = await etat(page);
    expect(saisiesHorsA(e), "saisies hors du champ « À »").toEqual([]);
    expect(await page.locator("#copilot-champ").inputValue()).toBe("");
  });

  test("« Votre session a expiré » (piège à focus) pendant le relevé : arrêt propre, « À » vidé, « Se connecter » jamais cliqué", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause que la fenêtre Letsignit : la recherche tapée reste dans « À ».
    const page = await ouvrir(context, sw, { releve: { type: "session", apres: 3 } }, "adresses");
    const avant = await demarrerReleve(page);
    await finReleve(page, avant);
    const e = await etat(page);
    expect(clicsInterdits(e)).toEqual([]);
    expect(page.url()).toContain("outlook.office.com");
    expect(saisiesHorsA(e)).toEqual([]);
    expect(e.aTexte, "texte laissé dans « À »").toBe("");
  });

  test("fenêtre Letsignit à piège asynchrone (façon tabster) : rien tapé ni cliqué dans la fenêtre", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { releve: { type: "letsignit", apres: 3, piege: "button", piegeMode: "async" } }, "adresses");
    const avant = await demarrerReleve(page);
    await finReleve(page, avant, 120000);
    const e = await etat(page);
    expect(saisiesHorsA(e)).toEqual([]);
    expect(clicsInterdits(e)).toEqual([]);
    expect(e.keysHorsA).toEqual([]);
    expect(e.aTexte).toBe("");
  });

  test("notification « Nouveau message de jean.dupont@… » pendant le relevé : son adresse n'est pas relevée, jamais cliquée", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { releve: { type: "toast", apres: 2 }, dureeToast: 20000 }, "adresses");
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    expect(msg).toContain("Terminé : 8 adresses");
    const e = await etat(page);
    expect(clicsInterdits(e)).toEqual([]);
    const base = String((await stock(sw))["base-outlook"] ?? "");
    expect(base).not.toContain("jean.dupont@ecole.fr");
    expect(base).not.toContain("support@ecole.fr"); // ni la liste des messages
    expect(e.aTexte).toBe("");
  });

  test("bandeau de mise à jour et rappel de réunion affichés pendant le relevé : jamais cliqués, relevé complet", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { maj: "bandeau", rappel: true }, "adresses");
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    expect(msg).toContain("Terminé : 8 adresses");
    const e = await etat(page);
    expect(clicsInterdits(e)).toEqual([]);
    expect(saisiesHorsA(e)).toEqual([]);
    expect(e.aTexte).toBe("");
  });

  test("« Démarrer » depuis la boîte de réception, Letsignit bloquant à l'ouverture du message : rien n'est tapé ni cliqué", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { letsignit: { quand: "ouverture", duree: null, piege: "button" } }, "adresses");
    await observer(page);
    const avant = await nbNotices(page, true);
    await ao(page).locator("#ao-run").click();
    await expect.poll(() => nbNotices(page, true), { timeout: 60000 }).toBeGreaterThan(avant);
    const e = await etat(page);
    expect(e.typed).toEqual([]);
    expect(e.aTexte).toBe("");
    expect(saisiesHorsA(e)).toEqual([]);
    expect(clicsInterdits(e)).toEqual([]);
  });

  test("« Démarrer » alors qu'une fenêtre (piège à focus) couvre le message ouvert : l'arrêt nomme la fenêtre", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : l'arrêt dit « Le nouveau message ne s'est pas ouvert avec le curseur dans le champ « À » » alors que le
    // message est ouvert et qu'une fenêtre d'Outlook le couvre (core/moteur.js:281-283 ; aucune détection de fenêtre
    // côté relevé).
    const page = await ouvrir(context, sw, {}, "adresses");
    await page.locator("#nouveau").click();
    await page.locator('[aria-label="À"][contenteditable="true"]').click();
    await page.evaluate(() => (window as any).__faux.letsignit({ duree: null, piege: "button" }));
    await observer(page);
    const avant = await nbNotices(page, true);
    await ao(page).locator("#ao-run").click();
    await expect.poll(() => nbNotices(page, true), { timeout: 60000 }).toBeGreaterThan(avant);
    const e = await etat(page);
    expect(e.typed).toEqual([]);
    expect(saisiesHorsA(e)).toEqual([]);
    expect((await ao(page).locator("#ao-notice").textContent()) ?? "").toMatch(/fenêtre|Letsignit/);
  });
});

// ============================================================ RELEVÉ — hors ligne, session, redirection, rechargement
test.describe("relevé : hors ligne, session expirée, redirection, rechargement", () => {
  test("connexion perdue pendant le relevé : pas de « Terminé », arrêt qui le dit, point de reprise gardé", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : une liste vide hors ligne est prise pour « aucune suggestion » : le relevé finit « Terminé : 2 adresses …
    // Toutes les pistes ont été explorées » (au lieu de 8) et efface le journal de reprise (core/moteur.js:541-548,
    // aucune vérification de connexion dans adapter-dom.js).
    const page = await ouvrir(context, sw, {}, "adresses");
    const avant = await demarrerReleve(page);
    await expect.poll(async () => (await etat(page)).typed.length, { timeout: 30000 }).toBeGreaterThanOrEqual(3);
    await context.setOffline(true);
    const msg = await finReleve(page, avant, 120000);
    await context.setOffline(false);
    expect(msg).not.toContain("Terminé");
    expect(msg).toMatch(/connexion|hors ligne/i);
    expect((await stock(sw))["reprise-outlook"], "point de reprise effacé").toBeTruthy();
  });

  test("hors ligne dès le départ : le relevé ne conclut pas « aucune adresse ne passe vos filtres »", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : même cause : 26 recherches vides puis « Terminé, mais aucune adresse ne passe vos filtres ».
    const page = await ouvrir(context, sw, {}, "adresses");
    await context.setOffline(true);
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    await context.setOffline(false);
    expect(msg).not.toContain("Terminé");
    expect(msg).toMatch(/connexion|hors ligne/i);
  });

  test("session expirée (fenêtre sans piège, l'annuaire ne répond plus) : pas de « Terminé » trompeur, reprise gardée", { tag: "@bug" }, async ({ context, sw }) => {
    // BUG : le relevé continue sous la fenêtre « Votre session a expiré », toutes les recherches reviennent vides, il
    // conclut « Terminé : 2 adresses … Toutes les pistes ont été explorées » et efface le point de reprise.
    const page = await ouvrir(context, sw, { releve: { type: "session", apres: 3, piege: false } }, "adresses");
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    expect(msg).not.toContain("Terminé");
    expect((await stock(sw))["reprise-outlook"], "point de reprise effacé").toBeTruthy();
  });

  test("redirection vers login.microsoftonline.com pendant le relevé : rien tapé sur la page de connexion, reprise sans retaper", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { releve: { type: "redirection", apres: 3 } }, "adresses");
    await demarrerReleve(page);
    await page.waitForURL(/login\.microsoftonline\.com/, { timeout: 30000 });
    await page.waitForTimeout(1500);
    expect(await page.evaluate(() => ({ panneaux: document.querySelectorAll("[data-pp-panneau],[data-testid='ao-panel']").length, saisies: (window as any).__saisies, valeur: (document.getElementById("identifiant") as HTMLInputElement).value })))
      .toEqual({ panneaux: 0, saisies: [], valeur: "" });
    await page.goto(URL_OUTLOOK);
    await pret(page);
    expect((await stock(sw))["reprise-outlook"]).toBeTruthy();
    await expect(ao(page).locator("#ao-run")).toHaveText("Reprendre");
    const faites = (await etat(page)).typed.slice();
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    expect(msg).toContain("Terminé : 8 adresses");
    const suite = (await etat(page)).typed.slice(faites.length);
    expect(suite.filter((q) => faites.slice(0, -1).includes(q))).toEqual([]); // rien de déjà fait n'est retapé
    expect((await etat(page)).aTexte).toBe("");
  });

  test("page rechargée (mise à jour d'Outlook) pendant le relevé : rien ne redémarre seul, reprise propre", async ({ context, sw }) => {
    const page = await ouvrir(context, sw, { releve: { type: "rechargement", apres: 3 } }, "adresses");
    await demarrerReleve(page);
    await expect.poll(async () => (await page.evaluate(() => (window as any).__faux?.etat().chargements ?? 0).catch(() => 0)), { timeout: 30000 }).toBeGreaterThanOrEqual(2);
    await pret(page);
    await page.waitForTimeout(2000);
    const e0 = await etat(page);
    const n0 = e0.typed.length;
    await expect(ao(page).locator("#ao-run")).toHaveText("Reprendre");
    expect(e0.redactionOuverte).toBe(false); // rien ne s'est rouvert ni tapé tout seul
    expect((await etat(page)).typed.length).toBe(n0);
    const avant = await demarrerReleve(page);
    const msg = await finReleve(page, avant, 120000);
    expect(msg).toContain("Terminé : 8 adresses");
    expect((await etat(page)).aTexte).toBe("");
  });
});
