// Audit Smart Mail — « Verrous entre relevé et envoi » : DEUX onglets Outlook, chacun avec son panneau.
// Banc et fausse page : tests/audit/verrous-annuaire/ (banc.ts, outlook.html). Aucune requête réelle.
// Le journal global des mails « envoyés » (banc.envois) compte les envois de TOUS les onglets, rechargements compris.
// Règles visées : E1, E2, E3, E7, M1, R6.
// Lancer : node node_modules/@playwright/test/cli.js test tests/audit-verrous-annuaire-onglets.spec.ts --workers=2 --reporter=line
import {
  test, expect, finEnvoi, pp, ouvrirOnglet, recharger, preparerEnvoi, lancerEnvoi, pauseEnvoi, attendreEnvois, demarrerReleve, attendreReleve,
  pauseReleve, etat, sonde, journal, lire, poser, liste, destinataires, ao, cfgPage, type Page, type Banc,
} from "./audit/verrous-annuaire/banc.ts";
import { readFileSync } from "node:fs";

const ZONES_MAIL = ["a", "cc", "cci", "objet", "corps"];
async function nouvelOnglet(banc: Banc, onglet: string, o: { rapide?: boolean; vue?: "adresses" | "envoi" } = {}) {
  return ouvrirOnglet(banc, { onglet, vue: o.vue ?? "envoi", rapide: o.rapide, page: await banc.context.newPage() });
}
function doublons(xs: string[]) {
  const n: Record<string, number> = {};
  for (const x of xs) n[x] = (n[x] ?? 0) + 1;
  return Object.entries(n).filter(([, k]) => k > 1).map(([x, k]) => `${x} ×${k}`);
}

test.describe("Deux onglets : la campagne partagée et le verrou de l'envoi", () => {
  test("VA-O-01 « Lancer l'envoi » dans un 2e onglet resté sur la préparation : refusé SANS toucher à la campagne en cours (rechargement : mêmes destinataires)", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B");       // ouvert avant la campagne : reste sur la préparation
    const LA = liste("a", 4), LB = liste("b", 6);
    await a.bringToFront();
    await preparerEnvoi(a, { dest: LA, objet: "Campagne A", pmin: 500, pmax: 500 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    await b.bringToFront();
    await preparerEnvoi(b, { dest: LB, objet: "Campagne B" });
    await lancerEnvoi(b);
    await expect(pp(b).locator("#pp-notice")).toContainText("Un envoi est déjà en cours dans un autre onglet Outlook.");
    expect.soft(await lire<string[]>(banc, "pp-liste"), "liste enregistrée de la campagne en cours, après le refus dans B").toEqual(LA);
    // A continue ; puis pause et rechargement de A (Cmd+R, mise à jour d'Outlook, plantage de l'onglet…).
    await a.bringToFront();
    await attendreEnvois(banc, 2);
    await pauseEnvoi(a);
    await recharger(a);
    await expect(pp(a).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    const reprise = await pp(a).locator("#pp-reprise-texte").textContent();
    await pp(a).locator("#pp-pmin").fill("0");
    await pp(a).locator("#pp-pmax").fill("0");
    await lancerEnvoi(a);
    await finEnvoi(a);
    // BUG : pp-go.onclick appelle moteur.creer() (écrit pp-liste et pp-campagne) AVANT que moteur.lancer() ne vérifie
    // le verrou des autres onglets. La campagne de A est écrasée par la liste de B ; après rechargement, A reprend
    // « Campagne A » sur la liste de B, à partir du curseur de A : des personnes de B reçoivent le mail de A.
    expect(destinataires(banc), `reprise annoncée : ${reprise}`).toEqual(LA);
  });

  test("VA-O-02 « Abandonner » (confirmé) dans un 2e onglet pendant l'envoi : plus aucun mail ne part, ou l'abandon est refusé", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const LA = liste("c", 4);
    await preparerEnvoi(a, { dest: LA, objet: "Campagne à abandonner", pmin: 300, pmax: 300 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B");       // ouvert pendant l'envoi : affiche « Envoi en cours de route »
    await expect(pp(b).locator("#pp-abandon")).toHaveText("Abandonner");
    b.once("dialog", (d) => void d.accept());       // « Abandonner cet envoi ? Les destinataires restants ne recevront rien. »
    await pp(b).locator("#pp-abandon").click();
    await page_wait(b, 1000);
    const noticeB = (await pp(b).locator("#pp-notice").textContent()) ?? "";
    const refuse = /autre onglet|en cours/i.test(noticeB) && (await pp(b).locator("#pp-abandon").isVisible());
    await page_wait(b, 8000);
    // BUG : « Abandonner » (pp-abandon.onclick) ne teste que moteur.enCours de SON onglet : il efface pp-campagne et
    // pp-liste et affiche « Prêt pour un nouvel envoi », mais l'onglet A, qui tient la campagne en mémoire, continue
    // d'envoyer aux « destinataires restants » qui « ne recevront rien ».
    expect(refuse || banc.envois.length === 1, `abandon accepté dans B (« ${noticeB} ») mais A a encore envoyé à : ${destinataires(banc).slice(1).join(", ")}`).toBe(true);
    await pauseEnvoi(a).catch(() => {});
  });

  test("VA-O-03 message corrigé dans un 2e onglet pendant l'envoi : appliqué aux mails restants (ou modification refusée)", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const LA = liste("t", 3);
    await preparerEnvoi(a, { dest: LA, objet: "Texte", texte: "Bonjour,\nTexte d'origine avec une faute.", pmin: 400, pmax: 400 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B");
    await expect(pp(b).locator("#pp-texte")).toHaveAttribute("contenteditable", "true");
    await attendreEnvois(banc, 2);
    const curseurA = 2;
    await pp(b).locator("#pp-texte").fill("Bonjour,\nTexte corrigé.");
    await page_wait(b, 1000); // enregistré (300 ms) : moteur.majMessage
    const c = await lire<{ curseur: number }>(banc, "pp-campagne");
    expect.soft(c.curseur, "progression enregistrée de la campagne après la correction faite dans B").toBeGreaterThanOrEqual(curseurA);
    await attendreEnvois(banc, 3);
    await finEnvoi(a);
    // BUG : B modifie SA copie (périmée) de la campagne et l'écrit (curseur de B compris) ; A envoie avec sa copie en
    // mémoire et réécrit la sienne au mail suivant : la correction est perdue sans un mot.
    expect(banc.envois[2].corps, "3e mail, parti après la correction faite dans B").toContain("Texte corrigé.");
  });

  test("VA-O-04 limite sur 24 h baissée dans un 2e onglet pendant l'envoi : respectée", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const LA = liste("l", 5);
    await preparerEnvoi(a, { dest: LA, objet: "Limite", pmin: 300, pmax: 300, max: 150 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B");
    await pp(b).locator("#pp-max").fill("2");
    await pp(b).locator("#pp-max").blur();
    await expect(pp(b).locator("#pp-estim")).toContainText("au plus 2 mails sur 24 h");
    await finEnvoi(a, 60000);
    // BUG : reglerRythme() écrit la copie de B ; A garde ses réglages en mémoire (c.reglages) et les réécrit : la
    // limite de 2 mails, affichée dans B, n'est pas appliquée.
    expect(banc.envois.length, `mails partis : ${destinataires(banc).join(", ")}`).toBeLessThanOrEqual(2);
  });

  test("VA-O-05 « Reprendre l'envoi » dans un 2e onglet pendant l'envoi : refusé, aucun doublon", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const LA = liste("r", 3);
    await preparerEnvoi(a, { dest: LA, objet: "Reprise ailleurs", pmin: 300, pmax: 300 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B");
    await expect(pp(b).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await lancerEnvoi(b);
    await expect(pp(b).locator("#pp-notice")).toContainText("Un envoi est déjà en cours dans un autre onglet Outlook.");
    await finEnvoi(a);
    expect(destinataires(banc)).toEqual(LA);
    expect((await sonde(b)).envoi).toBe(0);
  });

  test("VA-O-06 onglet de l'envoi gelé plus de 30 s (onglet caché bridé ou gelé par Chrome) : « Reprendre l'envoi » ailleurs ne double aucun envoi", async ({ banc }) => {
    test.setTimeout(150000);
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const LA = liste("g", 4);
    await preparerEnvoi(a, { dest: LA, objet: "Gel", pmin: 800, pmax: 800 }); // 8 s de pause (page rapide)
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B");
    await expect(pp(b).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    // Chrome gèle l'onglet A ou bride ses minuteries (en headless il ne le fait jamais : simulé par une pause du
    // débogueur sur A SEUL — plus aucune tâche de A ne tourne, B continue).
    const cdp = await banc.context.newCDPSession(a);
    await cdp.send("Debugger.enable");
    await cdp.send("Debugger.pause");
    await pp(b).locator("#pp-pmin").fill("0");
    await pp(b).locator("#pp-pmax").fill("0");
    await lancerEnvoi(b);
    await expect(pp(b).locator("#pp-notice")).toContainText("Un envoi est déjà en cours dans un autre onglet Outlook."); // verrou encore frais
    await expect.poll(async () => Date.now() - ((await lire<{ t: number }>(banc, "pp-actif"))?.t ?? 0), { timeout: 45000, intervals: [1000] }).toBeGreaterThan(30500);
    await lancerEnvoi(b); // l'utilisateur réessaie : le verrou de A est « périmé »
    await finEnvoi(b, 60000);
    await cdp.send("Debugger.resume");
    await cdp.send("Debugger.disable");
    await finEnvoi(a, 60000);
    // BUG : le verrou entre onglets (pp-actif) n'est rafraîchi que par setInterval(battement, 5000) de la page, que
    // Chrome bride (1 réveil/min après 5 min caché) ou gèle ; il n'est jamais revérifié dans la boucle, et la liste
    // « Déjà envoyés » en mémoire n'est relue qu'après chaque mail : A et B envoient au même destinataire.
    expect(doublons(destinataires(banc)), `envois : ${destinataires(banc).join(", ")}`).toEqual([]);
  });

  test("VA-O-07 « Déjà envoyés » importé dans un 2e onglet pendant la pause entre deux mails : la personne importée ne reçoit rien", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B");
    const LA = liste("d", 3);
    await a.bringToFront();
    await preparerEnvoi(a, { dest: LA, objet: "Déjà ailleurs", pmin: 400, pmax: 400 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    await pp(b).locator("#pp-deja-file").setInputFiles({ name: "deja-collegue.txt", mimeType: "text/plain", buffer: Buffer.from("d2@exemple.fr\n") });
    await expect(pp(b).locator("#pp-notice")).toContainText("1 adresse ajoutée à « Déjà envoyés »");
    expect(await lire<string[]>(banc, "pp-deja")).toContain("d2@exemple.fr");
    await finEnvoi(a);
    // BUG : la boucle d'envoi saute les « Déjà envoyés » avec sa liste EN MÉMOIRE (moteur.estDeja), relue seulement
    // au lancement et après chaque mail (ajouterDeja) : une adresse ajoutée ailleurs pendant la pause est envoyée.
    expect(destinataires(banc)).toEqual(["d1@exemple.fr", "d3@exemple.fr"]);
  });

  test("VA-O-08 « Déjà envoyés » vidé dans un 2e onglet pendant la pause : rien de déjà parti n'est renvoyé", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B");
    const LA = liste("v", 3);
    await a.bringToFront();
    await preparerEnvoi(a, { dest: LA, objet: "Vidage ailleurs", pmin: 300, pmax: 300 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    b.once("dialog", (d) => void d.accept());
    await expect(pp(b).locator("#pp-deja-vider")).toBeEnabled({ timeout: 5000 }).catch(() => {});
    if (await pp(b).locator("#pp-deja-vider").isEnabled()) await pp(b).locator("#pp-deja-vider").click();
    await finEnvoi(a);
    expect(destinataires(banc)).toEqual(LA);
  });
  test("VA-O-18 campagne en pause dans A, « Lancer l'envoi » dans un 2e onglet resté sur la préparation : la campagne de A n'est pas remplacée en silence", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B");
    const LA = liste("pa", 4), LB = liste("pb", 2);
    await a.bringToFront();
    await preparerEnvoi(a, { dest: LA, objet: "Campagne en pause", pmin: 600, pmax: 600 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    await pauseEnvoi(a);                       // 3 destinataires restants, « Reprendre l'envoi »
    await b.bringToFront();
    await preparerEnvoi(b, { dest: LB, objet: "Autre message", pmin: 0, pmax: 0 });
    await lancerEnvoi(b);
    await finEnvoi(b);
    const noticeB = await pp(b).locator("#pp-notice").textContent();
    await a.bringToFront();
    await recharger(a);
    const titre = await pp(a).locator("#pp-reprise-titre").textContent();
    const texte = await pp(a).locator("#pp-reprise-texte").textContent();
    // BUG : dans un seul onglet, remplacer une campagne en cours de route exige « Abandonner » (avec confirmation
    // « Les destinataires restants ne recevront rien »). Depuis un onglet resté sur la préparation, « Lancer l'envoi »
    // écrase pp-campagne/pp-liste sans rien demander : au rechargement, les 3 destinataires restants de A ont disparu.
    expect(`${titre} — ${texte}`, `B : « ${noticeB} »`).toContain("Envoi en cours de route");
    expect(texte).toContain("Objet : Campagne en pause");
  });
});

test.describe("Deux onglets : relevé et envoi chacun de son côté", () => {
  test("VA-O-09 relevé dans B pendant l'envoi de A : les mails de A sont exacts, aucune frappe de B dans A", async ({ banc }) => {
    test.setTimeout(150000);
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B", { rapide: false, vue: "adresses" });
    const LA = ["martin.dupont@exemple.fr", "marie.martin@exemple.fr", "emma.girard@exemple.fr"];
    await a.bringToFront();
    await preparerEnvoi(a, { dest: LA, objet: "Pendant un relevé", texte: "Bonjour,\nMessage exact.", pmin: 300, pmax: 300 });
    await lancerEnvoi(a);
    await b.bringToFront();
    await demarrerReleve(b);
    await attendreReleve(b);
    await finEnvoi(a);
    await pauseReleve(b);
    expect(banc.envois.map((m) => [m.onglet, m.a, m.cc, m.objet])).toEqual(LA.map((x) => ["A", [x], [], "Pendant un relevé"]));
    for (const m of banc.envois) expect(m.corps).toContain("Message exact.");
    const intrusA = (await journal(a)).frappes.filter((f) => !f.confiance && ZONES_MAIL.includes(f.zone));
    expect(intrusA).toEqual([]);
    const frappesB = (await journal(b)).frappes.filter((f) => !f.confiance);
    expect(frappesB.length).toBeGreaterThan(0);
    expect([...new Set(frappesB.map((f) => f.zone))]).toEqual(["a"]);
  });

  test("VA-O-10 deux relevés : le second onglet est refusé tant que le premier tourne", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", rapide: false, vue: "adresses" });
    const b = await nouvelOnglet(banc, "B", { rapide: false, vue: "adresses" });
    await a.bringToFront();
    await demarrerReleve(a);
    await attendreReleve(a);
    await b.bringToFront();
    await demarrerReleve(b);
    await expect(ao(b).locator("#ao-notice")).toContainText("semble déjà en cours");
    expect((await sonde(b)).releve).toBe(0);
    await a.bringToFront();
    await pauseReleve(a);
  });

  test("VA-O-11 envoi lancé dans B pendant un relevé dans A : mails de B exacts, le relevé de A ne touche à rien d'autre", async ({ banc }) => {
    test.setTimeout(150000);
    const a = await ouvrirOnglet(banc, { onglet: "A", rapide: false, vue: "adresses" });
    const b = await nouvelOnglet(banc, "B");
    await a.bringToFront();
    await demarrerReleve(a);
    await attendreReleve(a);
    await b.bringToFront();
    const LB = liste("x", 2);
    await preparerEnvoi(b, { dest: LB, objet: "Depuis B", pmin: 100, pmax: 100 });
    await lancerEnvoi(b);
    await finEnvoi(b);
    expect(banc.envois.map((m) => [m.onglet, m.a])).toEqual(LB.map((x) => ["B", [x]]));
    const intrusB = (await journal(b)).frappes.filter((f) => !f.confiance);
    expect(intrusB).toEqual([]);
    await a.bringToFront();
    const fa = (await journal(a)).frappes.filter((f) => !f.confiance);
    expect([...new Set(fa.map((f) => f.zone))]).toEqual(["a"]);
    await pauseReleve(a);
  });

  test("VA-O-12 nouvelles personnes relevées dans B pendant l'envoi à l'annuaire dans A : la liste de A ne bouge pas, « Nouvel envoi » ne propose qu'elles", async ({ banc }) => {
    test.setTimeout(150000);
    const BASE = ["martin.dupont@exemple.fr", "marie.martin@exemple.fr", "zoe.fabre@exemple.fr"];
    await poser(banc, { "base-outlook": BASE.join("\n") });
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await expect(pp(a).locator("#pp-nb")).toHaveText("3 adresses");
    await preparerEnvoi(a, { objet: "Annuaire", pmin: 500, pmax: 500 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B", { rapide: false, vue: "adresses" });
    await demarrerReleve(b);
    await attendreReleve(b);
    await expect.poll(async () => String((await lire<string>(banc, "mega-annuaire")) ?? "").split("\n").filter(Boolean).length, { timeout: 60000 }).toBeGreaterThan(8);
    await pauseReleve(b);
    await a.bringToFront();
    await finEnvoi(a);
    expect(destinataires(banc).sort()).toEqual([...BASE].sort());
    expect(await lire<string[]>(banc, "pp-liste")).toEqual([...BASE].sort());
    const annuaire = String(await lire<string>(banc, "base-outlook")).split("\n").filter(Boolean);
    await pp(a).locator("#pp-abandon").click(); // « Nouvel envoi » (campagne terminée : pas de confirmation)
    await expect(pp(a).locator("#pp-nb")).toHaveText(`${annuaire.filter((x) => !BASE.includes(x)).length} adresses`);
  });

  test("VA-O-13 case « Annuaire commun » décochée dans B pendant l'envoi à l'annuaire de A : la liste de A ne bouge pas", async ({ banc }) => {
    const BASE = liste("n", 4);
    await poser(banc, { "base-outlook": BASE.join("\n") });
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B");
    await a.bringToFront();
    await preparerEnvoi(a, { objet: "Annuaire décoché ailleurs", pmin: 200, pmax: 200 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    await pp(b).locator("#pp-annuaire-actif").uncheck();
    await expect(pp(b).locator("#pp-nb")).toHaveText("0 adresse");
    await finEnvoi(a);
    expect(destinataires(banc)).toEqual([...BASE].sort());
  });
});

test.describe("Deux onglets : copie de la campagne et des « Déjà envoyés » restée en mémoire", () => {
  test("VA-O-14 onglet A fermé pendant « Envoyer » (issue incertaine), « Reprendre l'envoi » dans B : le destinataire en cours n'est pas renvoyé", async ({ banc }) => {
    test.setTimeout(170000);
    const a = await ouvrirOnglet(banc, { onglet: "A", rapide: false, delai: 300, vue: "envoi" });
    const LA = liste("w", 3);
    await preparerEnvoi(a, { dest: LA, objet: "Fermeture en vol", pmin: 0, pmax: 0 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1, 90000);
    const b = await nouvelOnglet(banc, "B");        // ouvert après le 1er mail : sa copie de la campagne a curseur = 1
    await expect(pp(b).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await cfgPage(a, { fermetureEnvoi: 30000 });     // 2e mail : parti au clic, mais la rédaction tarde à se refermer
    await attendreEnvois(banc, 2, 90000);
    expect((await lire<{ resultats: Record<string, { s: string }> }>(banc, "pp-campagne")).resultats["1"]?.s).toBe("envoi");
    await a.close({ runBeforeUnload: false });       // onglet fermé (ou planté) pendant l'attente : rien n'a pu être noté
    await b.bringToFront();
    for (let i = 0; i < 40; i++) {                   // l'utilisateur relance depuis B (au besoin plusieurs fois)
      await lancerEnvoi(b);
      await finEnvoi(b, 90000);
      if (!/autre onglet/.test((await etat(b)).ppNotice)) break;
      await b.waitForTimeout(1000);
    }
    // BUG : B reprend avec SA copie de la campagne (chargée à l'ouverture) : moteur.lancer ne relit ni pp-campagne
    // (où w2 est noté « envoi », donc « à vérifier » après un rechargement) ni pp-liste ; w2, absent de « Déjà envoyés »
    // (noté seulement après la fermeture de la rédaction), reçoit un 2e mail.
    expect(doublons(destinataires(banc)), `envois : ${destinataires(banc).join(", ")}`).toEqual([]);
  });

  test("VA-O-15 onglet A fermé pendant sa pause (mail bien parti), « Reprendre l'envoi » dans B : personne n'est renvoyé", async ({ banc }) => {
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const LA = liste("y", 4);
    await preparerEnvoi(a, { dest: LA, objet: "Fermeture en pause", pmin: 1000, pmax: 1000 });
    await lancerEnvoi(a);
    await attendreEnvois(banc, 1);
    const b = await nouvelOnglet(banc, "B");
    await a.bringToFront();
    await expect.poll(async () => (await etat(a)).envoi).toBe(true);
    await a.close({ runBeforeUnload: false });
    await b.bringToFront();
    await pp(b).locator("#pp-pmin").fill("0");
    await pp(b).locator("#pp-pmax").fill("0");
    for (let i = 0; i < 40; i++) {
      await lancerEnvoi(b);
      await finEnvoi(b, 90000);
      if (!/autre onglet/.test((await etat(b)).ppNotice)) break;
      await b.waitForTimeout(1000);
    }
    expect(destinataires(banc)).toEqual(LA);
  });

  test("VA-O-16 « Télécharger » les « Déjà envoyés » dans un 2e onglet après un envoi : compte et fichier à jour (pas une liste périmée)", async ({ banc }) => {
    await poser(banc, { "pp-deja": ["ancien1@exemple.fr", "ancien2@exemple.fr"] });
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    const b = await nouvelOnglet(banc, "B");
    await expect(pp(b).locator("#pp-deja-nb")).toHaveText("2 adresses");
    await a.bringToFront();
    await preparerEnvoi(a, { dest: liste("z", 3), objet: "Fichier", pmin: 0, pmax: 0 });
    await lancerEnvoi(a);
    await finEnvoi(a);
    await expect(pp(a).locator("#pp-deja-nb")).toHaveText("5 adresses");
    await b.bringToFront();
    const nbB = await pp(b).locator("#pp-deja-nb").textContent();
    const dernierId = await banc.sw.evaluate(async () => Math.max(0, ...(await chrome.downloads.search({})).map((d) => d.id)));
    await pp(b).locator("#pp-deja-export").click();
    await expect(pp(b).locator("#pp-notice")).toContainText("publipostage-deja-envoyes.txt");
    // Playwright range les téléchargements sous un nom technique : on prend le 1er téléchargement terminé après le clic.
    let fichier = "";
    await expect.poll(async () => (fichier = await banc.sw.evaluate(async (id) => {
      const d = (await chrome.downloads.search({})).filter((x) => x.id > id && x.state === "complete").sort((x, y) => x.id - y.id)[0];
      return d ? d.filename : "";
    }, dernierId)), { timeout: 15000 }).not.toBe("");
    const lignes = fichier ? readFileSync(fichier, "utf8").split(/\r?\n/).filter(Boolean) : [];
    // BUG : B garde la liste « Déjà envoyés » chargée à son ouverture (aucune écoute de storage.onChanged, export sans
    // relecture) : il affiche 2 adresses et réécrit publipostage-deja-envoyes.txt (conflictAction « overwrite »)
    // avec cette liste périmée, effaçant les 3 personnes que A vient de servir.
    expect.soft(nbB, "compte affiché dans B").toBe("5 adresses");
    expect(lignes.sort(), `fichier téléchargé depuis B : ${fichier}`).toEqual(["ancien1@exemple.fr", "ancien2@exemple.fr", ...liste("z", 3)].sort());
  });

  test("VA-O-17 2e onglet ouvert après la fin d'un envoi à l'annuaire : « Nouvel envoi » n'y propose que les personnes jamais servies", async ({ banc }) => {
    await poser(banc, { "base-outlook": liste("s", 3).join("\n") });
    const a = await ouvrirOnglet(banc, { onglet: "A", vue: "envoi" });
    await preparerEnvoi(a, { objet: "Annuaire 1", pmin: 0, pmax: 0 });
    await lancerEnvoi(a);
    await finEnvoi(a);
    await poser(banc, { "base-outlook": liste("s", 5).join("\n") });
    const b = await nouvelOnglet(banc, "B");
    await expect(pp(b).locator("#pp-abandon")).toHaveText("Nouvel envoi");
    await pp(b).locator("#pp-abandon").click();
    await expect(pp(b).locator("#pp-nb")).toHaveText("2 adresses");
    await preparerEnvoi(b, { objet: "Annuaire 2", pmin: 0, pmax: 0 });
    await lancerEnvoi(b);
    await finEnvoi(b);
    expect(destinataires(banc)).toEqual(liste("s", 5));
  });
});

async function page_wait(p: Page, ms: number) { await p.waitForTimeout(ms); }
