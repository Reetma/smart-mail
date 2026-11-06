// Copie de la suite du Publipostage (navigateur/tests/envoi.spec.ts), lancée contre
// Smart Mail pour prouver que la fusion ne casse rien côté envoi.
// Bout en bout : l'extension réelle (navigateur/) chargée dans Chromium pilote une
// fausse page Outlook web servie à l'adresse https://outlook.office.com/mail/.
// Aucune requête ne sort : la page est fournie par page.route.
import { test as base, expect, chromium, type BrowserContext, type Page } from "@playwright/test";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";

const ICI = dirname(fileURLToPath(import.meta.url));
const EXT = join(ICI, "..", "extension");
const PAGE = readFileSync(join(ICI, "fausse-outlook.html"), "utf8");
const URL_OUTLOOK = "https://outlook.office.com/mail/";

const test = base.extend<{ context: BrowserContext; page: Page }>({
  // eslint-disable-next-line no-empty-pattern
  context: async ({}, use) => {
    const dir = mkdtempSync(join(tmpdir(), "publipostage-navigateur-"));
    const context = await chromium.launchPersistentContext(dir, {
      channel: "chromium",
      headless: true,
      args: [`--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`],
    });
    await context.route("https://outlook.office.com/**", (route) => route.fulfill({ status: 200, contentType: "text/html; charset=utf-8", body: PAGE }));
    await use(context);
    await context.close();
  },
  page: async ({ context }, use) => {
    // Smart Mail : onglet « ② Envoyer » affiché d'emblée.
    let [w] = context.serviceWorkers();
    if (!w) w = await context.waitForEvent("serviceworker");
    await w.evaluate(() => chrome.storage.local.set({ "mega-onglet": "envoi" }));
    const page = context.pages()[0] ?? (await context.newPage());
    await page.goto(URL_OUTLOOK);
    await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ timeout: 20000 });
    await use(page);
  },
});

const panneau = (page: Page) => page.locator("[data-pp-panneau]");
const envoyes = async (page: Page) => JSON.parse((await page.locator("html").getAttribute("data-envoyes")) ?? "[]") as Array<{ a: string[]; objet: string; corps: string; html: string }>;

// Un mail copié depuis Outlook ou Word : lien actif, gras, plus des pièges à retirer.
const MAIL_COLLE =
  '<!--StartFragment--><p class="MsoNormal">Bonjour,</p><p>&nbsp;</p><p>Inscrivez-vous <a href="https://exemple.fr/inscription" onclick="vol()">ici</a> avant <b>vendredi</b>.</p>' +
  '<p>Écrivez-nous : <a href="mailto:contact@exemple.fr">contact@exemple.fr</a></p><script>alert(1)</script><p><a href="javascript:alert(1)">piège</a></p><img src="x" onerror="alert(1)"><p>Cordialement</p>';

// Colle du HTML dans l'éditeur du panneau, comme Ctrl+A puis Ctrl+V (le message type est remplacé).
async function collerDansPanneau(page: Page, html: string): Promise<void> {
  await page.evaluate((h) => {
    const ed = document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-texte")!;
    ed.innerHTML = "";
    ed.focus();
    const dt = new DataTransfer();
    dt.setData("text/html", h);
    dt.setData("text/plain", "texte brut");
    ed.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  }, html);
}

function verifierLiens(html: string): void {
  expect(html).toContain('href="https://exemple.fr/inscription"');
  expect(html).toContain('href="mailto:contact@exemple.fr"');
  expect(html).not.toMatch(/javascript:|onclick|onerror|<script|<img/i);
  expect(html.indexOf("Bonjour")).toBeLessThan(html.indexOf("Signature"));
}

test("envoie un mail séparé à chaque destinataire, comme à la main", async ({ page }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill("alice@exemple.fr\nBob <bob@Exemple.FR>; alice@exemple.fr, carole@exemple.org");
  await expect(p.locator("#pp-nb")).toHaveText("3 adresses");
  await expect(p.locator("#pp-dest-info")).toHaveText("1 doublon retiré");
  await p.locator("#pp-objet").fill("Invitation");
  await p.locator("#pp-texte").fill("Bonjour,\n\nVoici l'invitation.\nÀ bientôt");
  await p.locator("#pp-pmin").fill("5");
  await p.locator("#pp-pmax").fill("6");
  await p.locator("#pp-go").click();

  await expect.poll(async () => (await envoyes(page)).length, { timeout: 60000 }).toBe(3);
  const liste = await envoyes(page);
  expect(liste.map((m) => m.a)).toEqual([["alice@exemple.fr"], ["bob@exemple.fr"], ["carole@exemple.org"]]);
  for (const m of liste) {
    expect(m.objet).toBe("Invitation");
    expect(m.corps).toContain("Bonjour,");
    expect(m.corps).toContain("Voici l'invitation.");
    expect(m.corps).toContain("À bientôt");
    expect(m.corps.indexOf("Bonjour")).toBeLessThan(m.corps.indexOf("Signature")); // texte avant la signature
  }
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 3 mail(s) envoyé(s).");
  await expect(p.locator("#pp-suivi-resume")).toHaveText("3 envoyé(s) sur 3");
  await expect(p.locator("#pp-go")).toBeDisabled();
});

test("pause puis reprise au destinataire suivant, sans renvoi", async ({ page }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill("a1@exemple.fr\na2@exemple.fr\na3@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("600");
  await p.locator("#pp-pmax").fill("600");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  await p.locator("#pp-pause").click();
  await expect(p.locator("#pp-notice")).toContainText("Envoi en pause");
  expect((await envoyes(page)).length).toBe(1);
  await expect(p.locator("#pp-go")).toHaveText("Reprendre l'envoi");

  // Rechargement : la campagne est conservée et reprend où elle en était.
  await page.reload();
  await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor();
  await expect(p.locator("#pp-go")).toHaveText("Reprendre l'envoi");
  await p.locator("#pp-pmin").fill("5");
  await p.locator("#pp-pmax").fill("5");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 60000 }).toBe(2);
  const tous = await envoyes(page);
  expect(tous.map((m) => m.a[0])).toEqual(["a2@exemple.fr", "a3@exemple.fr"]); // après rechargement, la page repart de zéro
  await expect(p.locator("#pp-suivi-resume")).toHaveText("3 envoyé(s) sur 3");
});

test("issue incertaine après « Envoyer » : arrêt, « à vérifier », jamais renvoyé", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => { (window as unknown as { mode: string }).mode = "bloque"; });
  await p.locator("#pp-dest").fill("x1@exemple.fr\nx2@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Ce destinataire ne sera pas renvoyé", { timeout: 60000 });
  await expect(p.locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 2, 1 à vérifier");
  expect((await envoyes(page)).length).toBe(0);
});

test("mail collé avec liens : nettoyé dans le panneau, liens actifs dans le mail envoyé", async ({ page }) => {
  const p = panneau(page);
  await collerDansPanneau(page, MAIL_COLLE);
  const editeur = await p.locator("#pp-texte").innerHTML();
  expect(editeur).toContain('<a href="https://exemple.fr/inscription">ici</a>');
  expect(editeur).not.toMatch(/javascript:|onclick|onerror|<script|<img|class=/i);
  await p.locator("#pp-dest").fill("lien@exemple.fr");
  await p.locator("#pp-objet").fill("Inscription");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 60000 }).toBe(1);
  const [m] = await envoyes(page);
  verifierLiens(m!.html);
  expect(m!.corps.replace(/\u00a0/g, " ")).toContain("Inscrivez-vous ici avant vendredi."); // espaces tapées : parfois insécables
  expect(m!.html).toMatch(/<a href="https:\/\/exemple\.fr\/inscription">ici<\/a>/);
});

test("mail collé : liens et gras conservés, et repli si Outlook ignore le collage", async ({ page }) => {
  const p = panneau(page);
  await collerDansPanneau(page, MAIL_COLLE);
  await p.locator("#pp-dest").fill("c1@exemple.fr\nc2@exemple.fr");
  await p.locator("#pp-objet").fill("Inscription");
  await p.locator("#pp-pmin").fill("300");
  await p.locator("#pp-pmax").fill("300");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 60000 }).toBe(1);
  // 2e mail : Outlook ignore les collages, l'extension insère elle-même (À et message).
  await page.evaluate(() => { const w = window as unknown as { collerIgnore: boolean; aCollerIgnore: boolean }; w.collerIgnore = true; w.aCollerIgnore = true; });
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 60000 }).toBe(2);
  for (const m of await envoyes(page)) {
    verifierLiens(m.html);
    expect(m.html).toContain("<b>vendredi</b>");
  }
});

test("gros fichier texte (50 000 adresses, une par ligne) et limite sur 24 h", async ({ page }) => {
  const p = panneau(page);
  const lignes = Array.from({ length: 50000 }, (_, i) => `personne${i}@exemple.fr`);
  lignes.splice(10, 0, "PERSONNE3@EXEMPLE.FR", "", "pas une adresse");
  await p.locator("#pp-file").setInputFiles({ name: "liste.txt", mimeType: "text/plain", buffer: Buffer.from(lignes.join("\n")) });
  await expect(p.locator("#pp-nb")).toHaveText(/^50\s000 adresses$/);
  await expect(p.locator("#pp-fichier-texte")).toHaveText(/liste\.txt : 50\s000 adresses/);
  await expect(p.locator("#pp-dest-info")).toHaveText("1 doublon retiré");
  await expect(p.locator("#pp-dest")).toHaveValue(""); // la liste n'est pas recopiée dans la zone de texte
  await p.locator("#pp-objet").fill("Grande liste");
  await p.locator("#pp-texte").fill("Bonjour");
  await p.locator("#pp-pmin").fill("5");
  await p.locator("#pp-pmax").fill("5");
  await p.locator("#pp-max").fill("2");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Limite de 2 mails sur 24 h atteinte", { timeout: 60000 });
  expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["personne0@exemple.fr", "personne1@exemple.fr"]);
  await expect(p.locator("#pp-suivi-resume")).toHaveText(/^2 envoyé\(s\) sur 50\s000$/);
  await expect(p.locator("#pp-go")).toHaveText("Reprendre l'envoi");
});

test("fenêtre Outlook ouverte par-dessus : arrêt avec son texte, rien n'est envoyé", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => {
    const d = document.createElement("div");
    d.setAttribute("role", "dialog");
    d.setAttribute("aria-modal", "true");
    d.innerHTML = "<p>Nouveautés d'Outlook : découvrez Copilot</p><button>OK</button>";
    document.body.appendChild(d);
    document.getElementById("zone")!.setAttribute("aria-hidden", "true");
  });
  await p.locator("#pp-dest").fill("bloque@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText(/Outlook affiche une fenêtre par-dessus depuis 45 s : « Nouveautés d'Outlook/, { timeout: 30000 });
  expect((await envoyes(page)).length).toBe(0);
  await expect(p.locator("#pp-go")).toHaveText("Reprendre l'envoi");
});

test("champs marqués aria-hidden par Outlook (sans fenêtre) : trouvés quand même", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => {
    // Outlook marque parfois toute l'application aria-hidden ; les champs restent affichés.
    new MutationObserver(() => { const c = document.getElementById("compose"); if (c && !c.hasAttribute("aria-hidden")) c.setAttribute("aria-hidden", "true"); })
      .observe(document.getElementById("zone")!, { childList: true });
  });
  await p.locator("#pp-dest").fill("cache@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
});

test("fenêtre passagère d'un complément (Letsignit) et fenêtre cachée hors écran : on attend, puis on envoie", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => {
    // Fenêtre d'exécution cachée, hors de l'écran, permanente : ne gêne pas.
    const cachee = document.createElement("div");
    cachee.setAttribute("role", "dialog");
    cachee.setAttribute("aria-modal", "true");
    cachee.style.cssText = "position:absolute;left:-10000px;top:0;width:300px;height:200px";
    cachee.textContent = "Letsignit runtime";
    document.body.appendChild(cachee);
    // À l'ouverture de la rédaction, « Letsignit Add-in » s'affiche 300 ms.
    new MutationObserver(() => {
      if (!document.getElementById("compose") || document.getElementById("lsg")) return;
      const d = document.createElement("div");
      d.id = "lsg";
      d.setAttribute("role", "dialog");
      d.setAttribute("aria-modal", "true");
      d.style.cssText = "position:fixed;left:30%;top:30%;width:300px;height:120px;background:#fff";
      d.textContent = "Letsignit Add-in";
      document.body.appendChild(d);
      setTimeout(() => d.remove(), 300);
    }).observe(document.getElementById("zone")!, { childList: true });
  });
  await p.locator("#pp-dest").fill("signature@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 1 mail(s) envoyé(s).");
});

test("adresses web en clair et bouton « Ajouter un lien » : liens cliquables dans le mail envoyé", async ({ page }) => {
  const p = panneau(page);
  // Texte simple collé (sans HTML) contenant des adresses web.
  await page.evaluate(() => {
    const ed = document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-texte")!;
    ed.innerHTML = "";
    ed.focus();
    const dt = new DataTransfer();
    dt.setData("text/plain", ["Programme : https://exemple.fr/programme.", "Site : www.exemple.org", "Inscription ici"].join(String.fromCharCode(10)));
    ed.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  const ed = p.locator("#pp-texte");
  await expect(ed.locator('a[href="https://exemple.fr/programme"]')).toHaveText("https://exemple.fr/programme");
  await expect(ed.locator('a[href="https://www.exemple.org"]')).toHaveText("www.exemple.org");
  await expect(p.locator("#pp-nb-liens")).toHaveText("2 liens cliquables → exemple.fr/programme, www.exemple.org");
  // Sélection du mot « ici », puis « Ajouter un lien ».
  await page.evaluate(() => {
    const e = document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-texte")!;
    const w = document.createTreeWalker(e, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = (n.nodeValue ?? "").indexOf("ici");
      if (i >= 0) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 3); const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r); break; }
    }
  });
  page.once("dialog", (d) => void d.accept("https://exemple.fr/inscription"));
  await p.locator("#pp-lien").click();
  await expect(ed.locator('a[href="https://exemple.fr/inscription"]')).toHaveText("ici");
  await expect(p.locator("#pp-nb-liens")).toHaveText("3 liens cliquables → exemple.fr/programme, www.exemple.org, exemple.fr/inscription");

  await p.locator("#pp-dest").fill("liens@exemple.fr");
  await p.locator("#pp-objet").fill("Programme et inscription");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  const [m] = await envoyes(page);
  expect(m!.a).toEqual(["liens@exemple.fr"]);
  expect(m!.objet).toBe("Programme et inscription");
  expect(m!.html).toContain('<a href="https://exemple.fr/programme">https://exemple.fr/programme</a>.');
  expect(m!.html).toContain('<a href="https://www.exemple.org">www.exemple.org</a>');
  expect(m!.html).toContain('<a href="https://exemple.fr/inscription">ici</a>');
});

test("garde-fou : si Outlook colle aussi autre chose (presse-papiers), rien n'est envoyé", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => {
    // Simule un Outlook qui ajoute le contenu du vrai presse-papiers en plus du message.
    document.addEventListener("paste", (e) => {
      const t = e.target as HTMLElement;
      if (!t.closest || !t.closest("#corps")) return;
      setTimeout(() => { document.getElementById("corps")!.insertAdjacentHTML("afterbegin", "<div>" + "Contenu inattendu du presse-papiers. ".repeat(10) + "</div>"); }, 0);
    }, true);
  });
  await p.locator("#pp-dest").fill("garde@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte court");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Le mail contient plus de texte que le message prévu", { timeout: 30000 });
  expect((await envoyes(page)).length).toBe(0);
});

test("réglages de pause respectés, y compris sous 5 s et modifiés pendant l'envoi", async ({ page }) => {
  const p = panneau(page);
  const pause = async () => Number(await panneau(page).getAttribute("data-pp-pause"));
  await p.locator("#pp-dest").fill("r1@exemple.fr\nr2@exemple.fr\nr3@exemple.fr");
  await p.locator("#pp-objet").fill("Rythme");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("2");
  await p.locator("#pp-pmax").fill("3");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBeGreaterThanOrEqual(1);
  await expect.poll(pause, { timeout: 10000 }).toBeGreaterThanOrEqual(2);
  expect(await pause()).toBeLessThanOrEqual(3); // et non 25 s par défaut
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(3);
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 3 mail(s) envoyé(s).");
});

test("pause raccourcie pendant l'envoi : appliquée tout de suite", async ({ page }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill("v1@exemple.fr\nv2@exemple.fr");
  await p.locator("#pp-objet").fill("Rythme");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("3000"); // 30 s en mode test : bien plus long que le test
  await p.locator("#pp-pmax").fill("3000");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  await expect(p.locator("#pp-live-etape")).toContainText("réglage : 3000 à 3000 s");
  const t0 = Date.now();
  await expect(p.locator("#pp-pmin")).toBeVisible(); // réglages visibles pendant l'envoi
  await p.locator("#pp-pmin").fill("1");
  await p.locator("#pp-pmax").fill("1");
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 8000 }).toBe(2);
  expect(Date.now() - t0).toBeLessThan(8000); // et non après les 30 s de la pause d'origine
});

test("valeurs de pause impossibles : la case montre la valeur réellement utilisée", async ({ page }) => {
  const p = panneau(page);
  await p.locator("#pp-pmin").fill("10");
  await p.locator("#pp-pmin").blur();
  await p.locator("#pp-pmax").fill("4"); // maximum plus petit que le minimum
  await p.locator("#pp-pmax").blur();
  await expect(p.locator("#pp-pmax")).toHaveValue("10");
  await p.locator("#pp-pmin").fill(""); // vide : on garde la valeur en cours
  await p.locator("#pp-pmin").blur();
  await expect(p.locator("#pp-pmin")).toHaveValue("10");
  await expect(p.locator("#pp-estim")).toContainText("Environ 22 s par mail");
});

// ------------------------------------------------------------ Cci, déjà envoyés, signature, importance
type Envoye = { a: string[]; cci: string[]; importance: string; objet: string; corps: string; html: string };
const envoyesCci = async (page: Page) => (await envoyes(page)) as unknown as Envoye[];
const adresses = (n: number, prefixe = "p") => Array.from({ length: n }, (_, i) => `${prefixe}${i}@exemple.fr`);

// Contenu du fichier « Déjà envoyés » écrit par le service de l'extension (Téléchargements).
async function fichierDeja(context: BrowserContext): Promise<string> {
  const sw = context.serviceWorkers()[0] ?? (await context.waitForEvent("serviceworker"));
  const dernier = () => sw.evaluate(async () => {
    type Telechargement = { filename: string; state: string };
    const c = (globalThis as unknown as { chrome: { downloads: { search: (q: object) => Promise<Telechargement[]> } } }).chrome;
    const d = (await c.downloads.search({ orderBy: ["-startTime"] })).find((x) => x.state === "complete");
    return d ? d.filename : "";
  });
  await expect.poll(dernier, { timeout: 10000 }).not.toBe("");
  return readFileSync(await dernier(), "utf8");
}

test("Cci : un mail pour 30 personnes, « À » vide, puis le paquet suivant", async ({ page }) => {
  const p = panneau(page);
  const liste = adresses(64).concat("alice@exemple.fr"); // alice : pastille au nom du contact
  await p.locator("#pp-dest").fill(liste.join("\n"));
  await p.locator("#pp-objet").fill("Invitation");
  await p.locator("#pp-texte").fill("Chers étudiants,\nInscrivez-vous.");
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  await p.locator("#pp-cci").check();
  await expect(p.locator("#pp-paquet")).toHaveValue("30");
  await expect(p.locator("#pp-max-unite")).toHaveText("destinataires");
  await expect(p.locator("#pp-estim")).toContainText("65 destinataires (3 mails en Cci)");
  await p.locator("#pp-max").fill("1000");
  await p.locator("#pp-go").click();

  await expect(p.locator("#pp-notice")).toContainText("Terminé : 65 destinataire(s) servi(s) en 3 mail(s).", { timeout: 60000 });
  const mails = await envoyesCci(page);
  expect(mails.map((m) => m.cci.length)).toEqual([30, 30, 5]);
  expect(mails.flatMap((m) => m.cci)).toEqual(liste);
  for (const m of mails) {
    expect(m.a).toEqual([]);
    expect(m.objet).toBe("Invitation");
    expect(m.corps).toContain("Inscrivez-vous.");
    expect(m.importance).toBe("normale");
  }
  await expect(p.locator("#pp-suivi-resume")).toHaveText("65 envoyé(s) sur 65");
  await expect(p.locator("#pp-deja-nb")).toHaveText("65 adresses");
});

test("Cci : la limite sur 24 h compte les destinataires", async ({ page }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill(adresses(50).join("\n"));
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  await p.locator("#pp-cci").check();
  await p.locator("#pp-max").fill("40");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Limite de 40 destinataires sur 24 h atteinte", { timeout: 60000 });
  expect((await envoyesCci(page)).map((m) => m.cci.length)).toEqual([30, 10]);
  await expect(p.locator("#pp-suivi-resume")).toHaveText("40 envoyé(s) sur 50");
  await expect(p.locator("#pp-go")).toHaveText("Reprendre l'envoi");
});

test("Cci : issue incertaine, tout le paquet « à vérifier » et jamais renvoyé", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => { (window as unknown as { mode: string }).mode = "bloque"; });
  await p.locator("#pp-dest").fill(adresses(5).join("\n"));
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-cci").check();
  await p.locator("#pp-paquet").fill("3");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Ces destinataires ne seront pas renvoyés", { timeout: 60000 });
  await expect(p.locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 5, 3 à vérifier");
  await expect(p.locator("#pp-deja-nb")).toHaveText("3 adresses");
  expect((await envoyes(page)).length).toBe(0);
});

test("déjà envoyés : fichier écrit, pas de renvoi au nouvel envoi, import et vidage", async ({ page, context }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill("d1@exemple.fr\nd2@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 2 mail(s) envoyé(s).", { timeout: 60000 });
  await expect(p.locator("#pp-deja-nb")).toHaveText("2 adresses");
  expect((await fichierDeja(context)).split(/\r?\n/).filter(Boolean)).toEqual(["d1@exemple.fr", "d2@exemple.fr"]);

  // Nouvel envoi avec les mêmes adresses et une nouvelle : seule la nouvelle part.
  await p.locator("#pp-abandon").click();
  await p.locator("#pp-dest").fill("d1@exemple.fr\nD2@exemple.fr\nd3@exemple.fr");
  await expect(p.locator("#pp-nb")).toHaveText("1 adresse");
  await expect(p.locator("#pp-dest-info")).toHaveText("2 déjà envoyées retirées");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 1 mail(s) envoyé(s).", { timeout: 60000 });
  expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["d1@exemple.fr", "d2@exemple.fr", "d3@exemple.fr"]);

  // Plus rien à envoyer à ces adresses.
  await p.locator("#pp-abandon").click();
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toHaveText("Toutes ces adresses ont déjà reçu le message (liste « Déjà envoyés »).");

  // Import d'un fichier (ajouté à la liste), puis vidage.
  await p.locator("#pp-deja-file").setInputFiles({ name: "deja.txt", mimeType: "text/plain", buffer: Buffer.from("z@exemple.fr\nd1@exemple.fr\n") });
  await expect(p.locator("#pp-notice")).toHaveText("1 adresse ajoutée à « Déjà envoyés » depuis deja.txt (1 était déjà dans la liste).");
  await expect(p.locator("#pp-deja-nb")).toHaveText("4 adresses");
  page.once("dialog", (d) => void d.accept());
  await p.locator("#pp-deja-vider").click();
  await expect(p.locator("#pp-deja-nb")).toHaveText("0 adresse");
  await expect(p.locator("#pp-nb")).toHaveText("3 adresses");
});

test("déjà envoyés importés pendant une pause : sautés à la reprise", async ({ page }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill("b1@exemple.fr\nb2@exemple.fr\nb3@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("600");
  await p.locator("#pp-pmax").fill("600");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  await p.locator("#pp-pause").click();
  await expect(p.locator("#pp-notice")).toContainText("Envoi en pause");
  await p.locator("#pp-deja-file").setInputFiles({ name: "autre-poste.txt", mimeType: "text/plain", buffer: Buffer.from("b2@exemple.fr\n") });
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 2 mail(s) envoyé(s). 1 adresse(s) déjà servie(s) avant : sautée(s).", { timeout: 60000 });
  expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["b1@exemple.fr", "b3@exemple.fr"]);
  await expect(p.locator("#pp-suivi-resume")).toHaveText("2 envoyé(s) sur 3, 1 déjà servi(s) avant");
});

// Signature IÉSEG : réseaux sociaux (images) | nom, classe, ville (texte) + logo (image).
const IMG = "data:image/gif;base64,R0lGODlhAQABAAAAACw=";
const SIGNATURE =
  `<table><tr><td><a href="https://facebook.com/ieseg"><img alt="facebook" src="${IMG}"></a><a href="https://linkedin.com/school/ieseg"><img alt="linkedin" src="${IMG}"></a></td>` +
  `<td style="border-left:2px solid #000"><div>DURANDEAU Paul</div><div>-BA1</div><div>Paris</div><img alt="IESEG" src="${IMG}"></td></tr></table>`;

test("signature : texte retiré, logo et réseaux gardés ; importance élevée par le ruban", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate((s) => { (window as unknown as { signature: string }).signature = s; }, SIGNATURE);
  await p.locator("#pp-dest").fill("s1@exemple.fr\ns2@exemple.fr");
  await p.locator("#pp-objet").fill("Certification IA");
  await p.locator("#pp-texte").fill("Chers étudiants,\nRendez-vous à\nParis\nle 27 octobre.");
  await p.locator("#pp-pmin").fill("0");
  await p.locator("#pp-pmax").fill("0");
  await p.locator("#pp-sans-signature").check();
  await p.locator("#pp-importance").check();
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 2 mail(s) envoyé(s).", { timeout: 60000 });
  for (const m of await envoyesCci(page)) {
    expect(m.importance).toBe("haute");
    expect(m.corps).not.toContain("DURANDEAU");
    expect(m.corps).not.toContain("BA1");
    expect(m.corps.match(/Paris/g)).toHaveLength(1); // celui du message, pas celui de la signature
    expect(m.corps).toContain("le 27 octobre.");
    for (const alt of ["facebook", "linkedin", "IESEG"]) expect(m.html).toContain(`alt="${alt}"`);
    expect(m.html).toContain('href="https://facebook.com/ieseg"');
  }
});

test("signature posée par un complément après le collage, et importance par « Plus d'options »", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate((s) => {
    const w = window as unknown as { signature: string; signatureSurCollage: string; importanceMenu: boolean };
    w.signature = "";
    w.signatureSurCollage = s;
    w.importanceMenu = true;
  }, SIGNATURE);
  await p.locator("#pp-dest").fill(adresses(4, "q").join("\n"));
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Bonjour");
  await p.locator("#pp-cci").check();
  await p.locator("#pp-sans-signature").check();
  await p.locator("#pp-importance").check();
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 4 destinataire(s) servi(s) en 1 mail(s).", { timeout: 60000 });
  const [m] = await envoyesCci(page);
  expect(m!.cci).toHaveLength(4);
  expect(m!.importance).toBe("haute");
  expect(m!.corps).not.toMatch(/DURANDEAU|BA1|Paris/);
  expect(m!.html).toContain('alt="IESEG"');
  await expect(page.locator('[role="menu"]')).toHaveCount(0); // menus refermés
});

test("importance : bouton du ruban sans état, repli sur « Plus d'options » vérifié en rouvrant le menu", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => { (window as unknown as { boutonSansEtat: boolean }).boutonSansEtat = true; });
  await p.locator("#pp-dest").fill("i1@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-importance").check();
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Terminé : 1 mail(s) envoyé(s).", { timeout: 60000 });
  expect((await envoyesCci(page))[0]!.importance).toBe("haute");
});

test("importance impossible à confirmer : arrêt, rien n'est envoyé", async ({ page }) => {
  const p = panneau(page);
  await page.evaluate(() => { const w = window as unknown as { boutonSansEtat: boolean; menuSansEtat: boolean }; w.boutonSansEtat = true; w.menuSansEtat = true; });
  await p.locator("#pp-dest").fill("i2@exemple.fr");
  await p.locator("#pp-objet").fill("Objet");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-importance").check();
  await p.locator("#pp-go").click();
  await expect(p.locator("#pp-notice")).toContainText("Outlook n'affiche pas « Importance haute » activée après le clic : rien n'est envoyé", { timeout: 60000 });
  expect((await envoyes(page)).length).toBe(0);
  await expect(p.locator("#pp-go")).toHaveText("Reprendre l'envoi");
});

// Place le curseur (ou une sélection) dans le texte du message.
async function selectionner(page: Page, mot: string, longueur: number): Promise<void> {
  await page.evaluate(([m, l]) => {
    const e = document.querySelector("[data-pp-panneau]")!.shadowRoot!.getElementById("pp-texte")!;
    const w = document.createTreeWalker(e, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const i = (n.nodeValue ?? "").indexOf(m as string);
      if (i >= 0) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + (l as number)); const s = window.getSelection()!; s.removeAllRanges(); s.addRange(r); break; }
    }
  }, [mot, longueur]);
}

test("lien vide resté d'un lien remplacé : retiré, le mail part avec le bon lien", async ({ page }) => {
  const p = panneau(page);
  await collerDansPanneau(page, '<div>Bonjour,</div><div><strong><span><a href="https://ancien.fr/ecole"></a><a href="https://nouveau.fr/ecole">OBTENEZ VOTRE CERTIFICATION</a></span></strong></div>');
  await expect(p.locator("#pp-nb-liens")).toHaveText("1 lien cliquable → nouveau.fr/ecole");
  await p.locator("#pp-dest").fill("vide@exemple.fr");
  await p.locator("#pp-objet").fill("Lien remplacé");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  const [m] = await envoyes(page);
  expect(m!.html).toContain('<a href="https://nouveau.fr/ecole">OBTENEZ VOTRE CERTIFICATION</a>');
  expect(m!.html).not.toContain("ancien.fr");
});

test("changer l'adresse d'un lien : curseur dans le lien, l'ancien disparaît en entier", async ({ page }) => {
  const p = panneau(page);
  await collerDansPanneau(page, '<div>Inscrivez-vous : <a href="https://ancien.fr/ecole">OBTENEZ VOTRE CERTIFICATION</a> avant vendredi.</div>');
  // Simple curseur dans le lien : l'adresse actuelle est proposée, puis remplacée.
  await selectionner(page, "VOTRE", 0);
  let proposee = "";
  page.once("dialog", (d) => { proposee = d.defaultValue(); void d.accept("https://nouveau.fr/school"); });
  await p.locator("#pp-lien").click();
  expect(proposee).toBe("https://ancien.fr/ecole");
  const ed = p.locator("#pp-texte");
  await expect(ed.locator('a[href="https://nouveau.fr/school"]')).toHaveText("OBTENEZ VOTRE CERTIFICATION");
  await expect(ed.locator("a")).toHaveCount(1);
  // Une partie du texte du lien sélectionnée : tout le lien change aussi, sans lien vide.
  await selectionner(page, "CERTIF", 6);
  page.once("dialog", (d) => void d.accept("https://troisieme.fr/ecole"));
  await p.locator("#pp-lien").click();
  await expect(ed.locator('a[href="https://troisieme.fr/ecole"]')).toHaveText("OBTENEZ VOTRE CERTIFICATION");
  await expect(ed.locator("a")).toHaveCount(1);
  await expect(p.locator("#pp-nb-liens")).toHaveText("1 lien cliquable → troisieme.fr/ecole");
  expect(await ed.innerHTML()).not.toMatch(/ancien|nouveau/);
});

test("message type déjà en place, modifiable, remis par « Message type », et envoyé tel quel", async ({ page }) => {
  const p = panneau(page);
  const ed = p.locator("#pp-texte");
  await expect(p.locator("#pp-objet")).toHaveValue("Certification IA - Avant 31 Oct.");
  await expect(ed).toContainText("Chers étudiants,");
  await expect(ed).toContainText("Reetma School");
  await expect(ed.locator('a[href="https://reetma.com/ecole"]')).toHaveCount(2);
  await expect(p.locator("#pp-nb-liens")).toHaveText("2 liens cliquables → reetma.com/ecole");
  // Modifiable, et la modification est gardée au rechargement.
  await p.locator("#pp-objet").fill("Autre objet");
  await ed.fill("Autre texte");
  await page.waitForTimeout(600); // brouillon enregistré
  await page.reload();
  await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor();
  await expect(p.locator("#pp-objet")).toHaveValue("Autre objet");
  await expect(ed).toHaveText("Autre texte");
  // « Message type » : confirmation, puis objet et texte remis.
  page.once("dialog", (d) => void d.accept());
  await p.locator("#pp-modele").click();
  await expect(p.locator("#pp-objet")).toHaveValue("Certification IA - Avant 31 Oct.");
  await expect(ed).toContainText("OBTENEZ VOTRE CERTIFICATION");
  await p.locator("#pp-dest").fill("modele@exemple.fr");
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 30000 }).toBe(1);
  const [m] = await envoyes(page);
  expect(m!.objet).toBe("Certification IA - Avant 31 Oct.");
  expect(m!.html).toContain('<a href="https://reetma.com/ecole">OBTENEZ VOTRE CERTIFICATION AETHER EN IA</a>');
  expect(m!.html).toContain('<a href="https://reetma.com/ecole">https://reetma.com/ecole</a>');
  expect(m!.corps).toContain("Cordialement,");
  expect(m!.corps.indexOf("Chers étudiants")).toBeLessThan(m!.corps.indexOf("Signature"));
});

test("onglet caché : minuteries de la page ralenties à une par minute, l'envoi avance quand même", async ({ page, context }) => {
  const p = panneau(page);
  await p.locator("#pp-dest").fill("fond1@exemple.fr\nfond2@exemple.fr\nfond3@exemple.fr");
  await p.locator("#pp-objet").fill("Arrière-plan");
  await p.locator("#pp-texte").fill("Texte");
  await p.locator("#pp-pmin").fill("2");
  await p.locator("#pp-pmax").fill("2");
  // Comme Chrome pour un onglet caché depuis 5 min : chaque minuterie des scripts
  // de l'extension (monde isolé, pas la page Outlook) attend une minute de plus.
  const cdp = await context.newCDPSession(page);
  const mondes: Array<{ id: number; name: string; auxData?: { type?: string } }> = [];
  cdp.on("Runtime.executionContextCreated", (e) => mondes.push(e.context));
  await cdp.send("Runtime.enable");
  const monde = mondes.find((c) => c.auxData?.type === "isolated" && c.name === "Smart Mail");
  expect(monde).toBeTruthy();
  await cdp.send("Runtime.evaluate", {
    contextId: monde!.id,
    expression: "(() => { const st = window.setTimeout; window.setTimeout = (f, ms, ...a) => st(f, (ms || 0) + 60000, ...a); })()",
  });
  const debut = Date.now();
  await p.locator("#pp-go").click();
  await expect.poll(async () => (await envoyes(page)).length, { timeout: 50000 }).toBe(3);
  expect(Date.now() - debut).toBeLessThan(50000); // sans les réveils de sw.js : plusieurs minutes par mail
  expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["fond1@exemple.fr", "fond2@exemple.fr", "fond3@exemple.fr"]);
});

test("version de l'extension affichée dans l'en-tête du panneau", async ({ page }) => {
  const version = (JSON.parse(readFileSync(join(EXT, "manifest.json"), "utf8")) as { version: string }).version;
  await expect(panneau(page).locator("#pp-version")).toHaveText(`v${version}`);
  await expect(panneau(page).locator("#pp-perime")).toBeHidden();
});
