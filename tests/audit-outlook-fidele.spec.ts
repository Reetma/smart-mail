// Audit « fidélité au vrai nouvel Outlook » (outlook.cloud.microsoft, FR, Letsignit).
// Smart Mail réelle contre tests/audit/outlook-fidele/outlook-reel.html (+ .js), qui
// reproduit ce que la fiche de référence sait du vrai Outlook, et des variantes
// PLAUSIBLES (signalées comme telles) pour ce qu'on ne sait pas. Un test qui échoue
// aujourd'hui exprime le comportement sûr attendu et porte un commentaire « // BUG : … ».
//   node node_modules/@playwright/test/cli.js test tests/audit-outlook-fidele.spec.ts --workers=2 --reporter=line
import {
  test, expect, ouvrir, recharger, pp, ao, onglet, envoyer, preparerEnvoi, lancerSansAttendre, finEnvoi, envoyes, etat, journal, suivi,
  clicsExtension, compact, texteDe, demarrerReleve, finReleve, baseReleve, type Journal, type Envoye,
} from "./audit/outlook-fidele/banc";
import type { Page } from "@playwright/test";

test.describe.configure({ mode: "parallel" });

// ------------------------------------------------------------ règles communes
// Commandes que l'envoi a le droit de toucher dans Outlook (E6).
const PERMIS_A = ["nouveau", "A", "OBJET", "CORPS", "envoyer"];
const PERMIS_CCI = ["nouveau", "bascule-cci", "CCI", "OBJET", "CORPS", "envoyer"];
const PERMIS_IMPORTANCE = ["importance-haute", "onglet:Message"];
const horsPermis = (clics: string[], ...permis: string[][]) => clics.filter((c) => !permis.some((p) => p.includes(c)));
// Champs où quelque chose a été écrit ou collé (hors panneau de l'extension).
const ecritures = (j: Journal, depuis = 0) => j.filter((e) => e.n >= depuis && (e.type === "saisie" || e.type === "collage") && e.cible !== "PANNEAU").map((e) => e.cible!);
// Pastilles créées (hors pré-remplissage d'une réponse ou d'un brouillon).
const pastillesCreees = (j: Journal, depuis = 0) => j.filter((e) => e.n >= depuis && e.type === "pastille" && e.source !== "pré-rempli" && e.source !== "test");
const adresses = (n: number, p = "etu") => Array.from({ length: n }, (_, i) => `${p}${i}@etudiants-test.fr`);
const SIGNATURE_TEXTES = ["Paul DURANDEAU", "Chargé de relations entreprises", "Campus de Paris", "+33 1 00 00 00 00"];
const SIGNATURE_IMAGES = ["Logo École Test", "LinkedIn", "Facebook", "Bannière portes ouvertes"];
const ANNUAIRE = [
  ["MARTIN Alice", "alice.martin@ecole-test.fr"], ["DUPONT Jean", "jean.dupont@ecole-test.fr"], ["BERNARD Camille", "camille.bernard@ecole-test.fr"],
  ["PETIT Thomas", "thomas.petit@ecole-test.fr"], ["ROBERT Paul", "paul.robert@ecole-test.fr"], ["RICHARD Emma", "emma.richard@ecole-test.fr"],
  ["DURAND Léa", "lea.durand@ecole-test.fr"], ["MOREAU Hugo", "hugo.moreau@ecole-test.fr"], ["SIMON Chloé", "chloe.simon@ecole-test.fr"],
  ["LAURENT Louis", "louis.laurent@ecole-test.fr"], ["MICHEL Manon", "manon.michel@ecole-test.fr"], ["GARCIA Lucas", "lucas.garcia@ecole-test.fr"],
  ["Alice PRUNIER", "alice.prunier@partenaire.example"], ["Arthur REMY", "arthur.remy@edu.example.com"], ["Miriam ABDER", "abder@example.com"],
  ["Scolarité Paris", "scolarite@ecole-test.fr"],
];
const ADRESSES_HORS_ANNUAIRE = ["noreply@email.apple.com", "learn@mails.microsoft.com", "secretariat@ecole-test.fr", "direction.campus@ecole-test.fr", "support@apple.example", "moi@ecole-test.fr"];

// Le message reçu doit être exactement le message prévu, suivi au plus de la signature.
async function verifierCorps(page: Page, m: Envoye, attenduHtml: string, signature: "avec" | "sans" | "texte-retire") {
  const attendu = await texteDe(page, attenduHtml);
  const recu = compact(m.texte);
  expect(recu.startsWith(attendu), `le mail commence par le message prévu\nreçu : ${recu.slice(0, 200)}\nprévu : ${attendu.slice(0, 200)}`).toBe(true);
  const reste = recu.slice(attendu.length);
  if (signature === "avec") expect(reste).toBe(compact(SIGNATURE_TEXTES.join("")));
  else expect(reste, "rien d'autre que le message").toBe("");
  if (signature !== "sans") for (const alt of SIGNATURE_IMAGES) expect(m.html).toContain(`alt="${alt}"`);
}

// ================================================================== ENVOI (À)
test.describe("envoi individuel (À)", () => {
  test("3 destinataires (2 de l'annuaire, pastille au NOM ; 1 externe) : un mail chacun, exact, rien d'autre de touché", async ({ context, worker, requetes }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr", "Jean.Dupont@ecole-test.fr", "etudiant.externe@gmail.example"], objet: "Invitation", texte: "Bonjour,\n\nVoici l'invitation.\nÀ bientôt" });
    expect(await envoyer(page)).toContain("Terminé : 3 mail(s) envoyé(s).");
    const m = await envoyes(page);
    expect(m.map((x) => x.a)).toEqual([["alice.martin@ecole-test.fr"], ["jean.dupont@ecole-test.fr"], ["etudiant.externe@gmail.example"]]);
    for (const x of m) {
      expect([x.cc, x.cci, x.libres]).toEqual([[], [], []]);
      expect(x.objet).toBe("Invitation");
      expect(x.importance).toBe("normale");
      expect(compact(x.texte)).toBe(compact("Bonjour,Voici l'invitation.À bientôt"));
    }
    expect(horsPermis(await clicsExtension(page), PERMIS_A)).toEqual([]);
    expect(new Set(ecritures(await journal(page)))).toEqual(new Set(["A", "OBJET", "CORPS"]));
    const e = await etat(page);
    expect([e.composes.length, e.carnet, e.modales.length]).toEqual([0, 0, 0]);
    expect(requetes).toEqual([]);
  });

  test("message type (2 liens reetma.com, gras, liens marqués data-auth par Outlook) : objet, texte et liens exacts", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"] }); // objet et texte : le message type en place
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const [m] = await envoyes(page);
    expect(m!.objet).toBe("Certification IA - Avant 31 Oct.");
    const modele = await pp(page).evaluate((h) => (h as HTMLElement).shadowRoot!.getElementById("pp-texte")!.innerHTML);
    await verifierCorps(page, m!, modele, "sans");
    expect((m!.html.match(/href="https:\/\/reetma\.com\/ecole"/g) ?? []).length).toBe(2);
    expect(m!.html).toMatch(/<strong>.*OBTENEZ VOTRE CERTIFICATION AETHER EN IA.*<\/strong>/);
  });

  test("Letsignit pose sa signature à l'ouverture (fenêtre « Letsignit Add-in » passagère) : message AVANT la signature, rien de perdu", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { delai: 300, modale: { apres: 100, duree: 1500, piege: true } } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr", "jean.dupont@ecole-test.fr"] });
    expect(await envoyer(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    const modele = await pp(page).evaluate((h) => (h as HTMLElement).shadowRoot!.getElementById("pp-texte")!.innerHTML);
    for (const m of await envoyes(page)) await verifierCorps(page, m, modele, "avec");
    expect(horsPermis(await clicsExtension(page), PERMIS_A)).toEqual([]);
  });

  test("Letsignit pose sa signature APRÈS le collage du message : le message reste entier et en tête", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { declencheur: "collage-corps", delai: 0 } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"] });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const modele = await pp(page).evaluate((h) => (h as HTMLElement).shadowRoot!.getElementById("pp-texte")!.innerHTML);
    await verifierCorps(page, (await envoyes(page))[0]!, modele, "avec");
  });

  test("fenêtre « Letsignit Add-in » qui ne se ferme pas (connexion au complément) : arrêt clair, rien n'est envoyé, aucun bouton de la fenêtre cliqué", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { html: false, modale: { apres: 50, duree: null, piege: true } } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr", "jean.dupont@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    const r = await envoyer(page);
    expect(r).toMatch(/Arrêt avant l'envoi à alice\.martin@ecole-test\.fr \(rien n'est parti\) : Outlook affiche une fenêtre par-dessus depuis 45 s : « Letsignit Add-in/);
    expect(await envoyes(page)).toEqual([]);
    expect((await journal(page)).filter((e) => e.type === "dialogue-bouton")).toEqual([]);
    await expect(pp(page).locator("#pp-go")).toHaveText("Reprendre l'envoi");
  });

  test("une rédaction de l'utilisateur est ouverte (brouillon avec destinataire, objet, texte) : arrêt sans envoi, brouillon intact", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { brouillon: { a: ["paul.robert@ecole-test.fr"], objet: "Mon brouillon", corps: "Paul, voici mes notes." } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Campagne", texte: "Texte" });
    expect(await envoyer(page)).toContain("Un message est déjà ouvert en rédaction");
    expect(await envoyes(page)).toEqual([]);
    const [c] = (await etat(page)).composes;
    expect(c!.a.map((x) => x.adresse)).toEqual(["paul.robert@ecole-test.fr"]);
    expect([c!.objet, compact(c!.corps), c!.libres.A]).toEqual(["Mon brouillon", compact("Paul, voici mes notes."), ""]);
  });

  test("une réponse en ligne est ouverte (À pré-rempli, pas d'objet visible) : arrêt sans envoi, réponse intacte", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { reponse: { depuis: 1 } });
    await preparerEnvoi(page, { dest: ["jean.dupont@ecole-test.fr"], objet: "Campagne", texte: "Texte" });
    expect(await envoyer(page)).toContain("Un message est déjà ouvert en rédaction");
    expect(await envoyes(page)).toEqual([]);
    const [c] = (await etat(page)).composes;
    expect(c!.a.map((x) => x.adresse)).toEqual(["alice.martin@ecole-test.fr"]);
    expect(c!.libres.A).toBe("");
    expect(c!.objet).toBe("RE: Réunion de rentrée");
  });

  test("le brouillon de l'utilisateur est dans un onglet masqué : un NOUVEAU message est ouvert, le brouillon n'est pas touché", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { brouillon: { a: ["paul.robert@ecole-test.fr"], objet: "Mon brouillon", corps: "Notes." } });
    await page.locator('[data-ident="liste:0"] .de').click(); // l'utilisateur lit un autre message : le brouillon passe en onglet
    expect((await etat(page)).composes.map((c) => c.visible)).toEqual([false]);
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Campagne", texte: "Texte" });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const [m] = await envoyes(page);
    expect([m!.a, m!.type, m!.objet]).toEqual([["alice.martin@ecole-test.fr"], "nouveau", "Campagne"]);
    const [c] = (await etat(page)).composes;
    expect([c!.type, c!.a.map((x) => x.adresse), c!.objet, compact(c!.corps)]).toEqual(["brouillon", ["paul.robert@ecole-test.fr"], "Mon brouillon", "Notes."]);
  });

  test("Outlook ignore le collage synthétique dans « À » (repli prévu par le code) : l'adresse tapée part telle quelle", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { collageIgnore: ["A"] });
    await preparerEnvoi(page, { dest: ["etudiant.externe@gmail.example", "alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    expect(await envoyer(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    expect((await envoyes(page)).map((m) => m.a)).toEqual([["etudiant.externe@gmail.example"], ["alice.martin@ecole-test.fr"]]);
  });

  test("variante plausible : collage ignoré ET « ; » qui valide la meilleure suggestion : jamais d'envoi à quelqu'un d'autre que prévu (E1)", async ({ context, worker }) => {
    // alice@ecole-test.fr n'est pas dans l'annuaire ; « MARTIN Alice <alice.martin@ecole-test.fr> »
    // répond à « alice », « ecole », « test », « fr » : c'est la suggestion qu'Outlook valide.
    const page = await ouvrir(context, worker, { collageIgnore: ["A"], pointVirgule: "suggestion" });
    await preparerEnvoi(page, { dest: ["alice@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    const r = await envoyer(page);
    const m = await envoyes(page);
    // BUG : remplirA (publipostage/outlook.js:367-372, aValidee) accepte n'importe quelle pastille
    // nouvelle sans vérifier qu'elle porte l'adresse prévue : la pastille « MARTIN Alice »
    // (alice.martin@…) est prise pour alice@ecole-test.fr et le mail part chez la mauvaise personne.
    expect(m.map((x) => x.a), "aucun mail chez un autre destinataire que prévu : " + r).toEqual([]);
  });

  test("adresses inhabituelles (apostrophe, +, majuscules dans le nom) : la pastille montre l'adresse, le mail part exactement à elle", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["o'brien+promo@Etudiants-Test.fr", "Zoe.Durand@etudiants-test.fr"], objet: "Objet", texte: "Texte" });
    expect(await envoyer(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    expect((await envoyes(page)).map((m) => m.a)).toEqual([["o'brien+promo@etudiants-test.fr"], ["Zoe.Durand@etudiants-test.fr"]]);
  });

  test("message riche collé (gras, liste, liens) : identique dans le mail, liens intacts", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    const html = '<p>Bonjour,</p><p>Inscrivez-vous <a href="https://exemple.fr/inscription">ici</a> avant <b>vendredi</b>.</p><ul><li>Atelier IA</li><li>Conférence</li></ul><p>Contact : <a href="mailto:contact@exemple.fr">contact@exemple.fr</a></p>';
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Inscription", html });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const [m] = await envoyes(page);
    await verifierCorps(page, m!, html, "sans");
    expect(m!.html).toMatch(/<a href="https:\/\/exemple\.fr\/inscription"[^>]*>ici<\/a>/);
    expect(m!.html).toMatch(/<a href="mailto:contact@exemple\.fr"[^>]*>contact@exemple\.fr<\/a>/);
    expect(m!.html).toContain("<b>vendredi</b>");
  });

  test("Outlook ignore le collage HTML dans le corps (repli insertHTML prévu par le code) : texte et liens exacts", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { collageIgnore: ["CORPS"] });
    const html = '<p>Bonjour,</p><p>Le programme : <a href="https://exemple.fr/programme">programme</a>.</p>';
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Programme", html });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const [m] = await envoyes(page);
    await verifierCorps(page, m!, html, "sans");
    expect(m!.html).toContain('href="https://exemple.fr/programme"');
  });

  test("bulle « Nouveautés d'Outlook » non modale ouverte : elle ne bloque pas et n'est jamais cliquée", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await page.evaluate(() => {
      const d = document.createElement("div");
      d.setAttribute("role", "dialog"); d.setAttribute("aria-label", "Nouveautés d'Outlook"); d.setAttribute("data-ident", "bulle");
      d.style.cssText = "position:fixed;left:240px;top:140px;width:260px;height:120px;background:#fff;border:1px solid #888;z-index:50";
      d.innerHTML = "<p>Découvrez Copilot dans Outlook.</p><button data-ident='bulle:ok'>Compris</button>";
      document.body.appendChild(d);
    });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    expect(horsPermis(await clicsExtension(page), PERMIS_A)).toEqual([]);
  });

  test("double clic sur « Lancer l'envoi » : un seul mail par destinataire (E2)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["d1@ecole-test.fr", "d2@ecole-test.fr", "d3@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 50, pmax: 50 });
    const avant = (await suivi(page, "__ppSuivi")).fins;
    await pp(page).locator("#pp-go").dblclick();
    await finEnvoi(page, avant);
    await page.waitForTimeout(1500);
    expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["d1@ecole-test.fr", "d2@ecole-test.fr", "d3@ecole-test.fr"]);
  });

  test("deux onglets Outlook ouverts : le second refuse de lancer pendant que le premier envoie, et ne renvoie rien ensuite (E2)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    const page2 = await context.newPage();
    await page2.goto("https://outlook.cloud.microsoft/mail/inbox");
    await page2.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached" });
    await page.bringToFront();
    await preparerEnvoi(page, { dest: ["t1@ecole-test.fr", "t2@ecole-test.fr", "t3@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 150, pmax: 150 });
    const avant = (await suivi(page, "__ppSuivi")).fins;
    await lancerSansAttendre(page);
    await expect.poll(async () => (await envoyes(page)).length, { timeout: 20000 }).toBe(1);
    await page2.bringToFront();
    await page2.reload();
    await page2.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached" });
    await pp(page2).locator("#pp-go").click();
    await expect(pp(page2).locator("#pp-notice")).toContainText("Un envoi est déjà en cours dans un autre onglet Outlook.");
    await page.bringToFront();
    expect(await finEnvoi(page, avant, 60000)).toContain("Terminé : 3 mail(s) envoyé(s).");
    // L'autre onglet, resté sur l'état d'avant, « reprend » : personne ne reçoit deux fois.
    await page2.bringToFront();
    await pp(page2).locator("#pp-go").click().catch(() => {});
    await page2.waitForTimeout(3000);
    expect(await envoyes(page2)).toEqual([]);
    expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["t1@ecole-test.fr", "t2@ecole-test.fr", "t3@ecole-test.fr"]);
  });
});

// ================================================================== ÉTATS INATTENDUS
test.describe("états inattendus d'Outlook (E5, E6)", () => {
  test("fenêtre « destinataires externes » à l'envoi : jamais « Envoyer quand même », « à vérifier », jamais renvoyé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { envoiDialogue: "externes" });
    await preparerEnvoi(page, { dest: ["externe1@gmail.example", "externe2@gmail.example"], objet: "Objet", texte: "Texte" });
    const r = await envoyer(page);
    expect(r).toContain("externe1@gmail.example : Après « Envoyer », la rédaction est restée ouverte");
    expect(r).toContain("Destinataires externes");
    expect(await envoyes(page)).toEqual([]);
    expect((await journal(page)).filter((e) => e.type === "dialogue-bouton")).toEqual([]);
    await expect(pp(page).locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 2, 1 à vérifier");
    // Reprendre : la fenêtre est toujours là, rien ne part, rien n'est cliqué.
    expect(await envoyer(page)).toMatch(/Arrêt avant l'envoi à externe2@gmail\.example \(rien n'est parti\)/);
    expect(await envoyes(page)).toEqual([]);
    expect((await journal(page)).filter((e) => e.type === "dialogue-bouton")).toEqual([]);
  });

  test("étiquette de sensibilité obligatoire à l'envoi : jamais « OK », « à vérifier », rien d'envoyé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { envoiDialogue: "sensibilite" });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    expect(await envoyer(page)).toContain("Étiquette de sensibilité requise");
    expect(await envoyes(page)).toEqual([]);
    expect((await journal(page)).filter((e) => e.type === "dialogue-bouton")).toEqual([]);
    expect(horsPermis(await clicsExtension(page), PERMIS_A)).toEqual([]);
  });

  test("session expirée (fenêtre « Actualiser ») pendant la pause entre deux mails : arrêt clair, rien de plus envoyé, rien cliqué", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["s1@ecole-test.fr", "s2@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 300, pmax: 300 }); // 3 s en mode accéléré
    const avant = (await suivi(page, "__ppSuivi")).fins;
    await lancerSansAttendre(page);
    await expect.poll(async () => (await envoyes(page)).length, { timeout: 20000 }).toBe(1);
    await page.evaluate(() => (window as unknown as { __fausse: { sessionExpiree: () => void } }).__fausse.sessionExpiree());
    const r = await finEnvoi(page, avant);
    expect(r).toMatch(/Arrêt avant l'envoi à s2@ecole-test\.fr \(rien n'est parti\) : Outlook affiche une fenêtre par-dessus depuis 45 s : « Session expirée/);
    expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["s1@ecole-test.fr"]);
    expect((await journal(page)).filter((e) => e.type === "dialogue-bouton")).toEqual([]);
  });

  test("redirection vers la page de connexion pendant la pause : au retour, rien ne repart seul, puis reprise sans doublon (E2, E5)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["c1@ecole-test.fr", "c2@ecole-test.fr", "c3@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 500, pmax: 500 });
    await lancerSansAttendre(page);
    await expect.poll(async () => (await envoyes(page)).length, { timeout: 20000 }).toBe(1);
    await page.evaluate(() => (window as unknown as { __fausse: { redirigerConnexion: () => void } }).__fausse.redirigerConnexion());
    await page.waitForURL(/login\.microsoftonline\.com/);
    await page.waitForTimeout(1500);
    await page.goto("https://outlook.cloud.microsoft/mail/inbox"); // l'utilisateur se reconnecte
    await page.locator('[data-pp-panneau][data-pp-pret="1"]').waitFor({ state: "attached" });
    await page.waitForTimeout(3000);
    expect(await envoyes(page)).toEqual([]); // nouvelle page : rien n'est reparti tout seul
    await expect(pp(page).locator("#pp-go")).toHaveText("Reprendre l'envoi");
    await expect(pp(page).locator("#pp-suivi-resume")).toHaveText("1 envoyé(s) sur 3");
  });

  test("hors connexion pendant la campagne : arrêt SANS envoi avec un message clair (E5), pas de mails en file qui partent d'un coup au retour (E7)", async ({ context, worker }) => {
    // Variante plausible du nouvel Outlook hors connexion : la rédaction se ferme et le message
    // attend dans la « Boîte d'envoi » ; Outlook l'envoie dès que la connexion revient.
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["h1@ecole-test.fr", "h2@ecole-test.fr", "h3@ecole-test.fr", "h4@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 100, pmax: 100 }); // 1 s
    const avant = (await suivi(page, "__ppSuivi")).fins;
    await lancerSansAttendre(page);
    await expect.poll(async () => (await envoyes(page)).length, { timeout: 20000 }).toBe(1);
    await context.setOffline(true);
    await expect(page.locator('[data-ident="hors-ligne"]')).toBeVisible();
    const r = await finEnvoi(page, avant);
    const boite = Number(await page.locator("html").getAttribute("data-boite-envoi"));
    await context.setOffline(false);
    await page.waitForTimeout(300);
    const m = await envoyes(page);
    // BUG : rien ne vérifie la connexion (aucun navigator.onLine / bandeau « hors connexion »
    // dans publipostage/outlook.js ni moteur.js) : hors ligne, chaque mail est compté « envoyé »
    // dès que la rédaction se ferme, la campagne continue, et au retour de la connexion
    // Outlook expédie toute la Boîte d'envoi d'un coup, sans les pauses réglées.
    expect(boite, "aucun mail laissé en Boîte d'envoi : " + r).toBe(0);
    expect(m.filter((x) => x.depuisBoiteEnvoi).map((x) => x.a[0])).toEqual([]);
    expect(r).toMatch(/connexion|hors ligne/i);
  });
});

test.describe("états inattendus d'Outlook (suite)", () => {
  test("variante plausible : une fenêtre d'Outlook avec un champ de saisie (enquête) surgit au moment de l'objet : rien n'est tapé dedans, rien n'est envoyé de travers", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { modaleSurObjet: { duree: 2000 } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Certification IA", texte: "Texte" });
    const r = await envoyer(page);
    const j = await journal(page);
    const dansLaFenetre = j.filter((e) => (e.type === "saisie" || e.type === "collage") && String(e.cible).startsWith("dialogue:"));
    const m = await envoyes(page);
    for (const x of m) expect(x.objet).toBe("Certification IA");
    // BUG : remplirObjet → coller → insererTexte (publipostage/outlook.js:209-213, 339-353) tape
    // l'objet par execCommand("insertText") dans l'élément qui a le focus, sans vérifier que c'est
    // l'objet : la fenêtre (piège à focus) l'a pris, l'objet du mail part dans son champ.
    expect(dansLaFenetre.map((e) => e.cible + " ← " + e.data), r).toEqual([]);
  });

  test("objet de plus de 255 caractères (maxlength du vrai champ « Objet ») : jamais d'envoi avec un objet tronqué (E3)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    const objet = "Certification IA – " + "inscriptions ouvertes ".repeat(14).trim(); // > 255
    expect(objet.length).toBeGreaterThan(255);
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet, texte: "Texte" });
    const r = await envoyer(page);
    for (const m of await envoyes(page)) expect(m.objet, r).toBe(objet);
  });

  test("pas de bouton « Nouveau message » (l'utilisateur est dans le Calendrier) : arrêt clair, rien n'est tapé ni envoyé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await page.evaluate(() => { document.getElementById("commandes")!.innerHTML = "<button type='button' data-ident='calendrier:nouvel-evenement'>Nouvel événement</button>"; });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    expect(await envoyer(page)).toContain("Bouton « Nouveau message » introuvable");
    expect(await envoyes(page)).toEqual([]);
    expect(ecritures(await journal(page))).toEqual([]);
    expect(await clicsExtension(page)).toEqual([]);
  });
});

// ================================================================== CCI
test.describe("envoi en Cci (paquets)", () => {
  test("7 destinataires en paquets de 3 : « Cci » affiché par sa bascule, JAMAIS le libellé « Cci » (carnet d'adresses), À et Cc vides", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    const liste = adresses(6).concat("alice.martin@ecole-test.fr"); // la dernière : pastille au nom
    await preparerEnvoi(page, { dest: liste, objet: "Invitation", texte: "Chers étudiants,\nInscrivez-vous.", cci: true, paquet: 3, max: 1000 });
    expect(await envoyer(page)).toContain("Terminé : 7 destinataire(s) servi(s) en 3 mail(s).");
    const m = await envoyes(page);
    expect(m.map((x) => x.cci)).toEqual([liste.slice(0, 3), liste.slice(3, 6), liste.slice(6)]);
    for (const x of m) { expect([x.a, x.cc, x.libres]).toEqual([[], [], []]); expect(x.objet).toBe("Invitation"); }
    expect(horsPermis(await clicsExtension(page), PERMIS_CCI)).toEqual([]);
    expect((await etat(page)).carnet).toBe(0);
    expect(new Set(ecritures(await journal(page)))).toEqual(new Set(["CCI", "OBJET", "CORPS"]));
  });

  test("« Toujours afficher Cci » : le champ est déjà là, aucune bascule cliquée, 30 adresses d'un coup", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { cciVisible: true });
    const liste = adresses(30);
    await preparerEnvoi(page, { dest: liste, objet: "Objet", texte: "Texte", cci: true, paquet: 30, max: 1000 });
    expect(await envoyer(page)).toContain("Terminé : 30 destinataire(s) servi(s) en 1 mail(s).");
    const [m] = await envoyes(page);
    expect([m!.cci, m!.a, m!.cc]).toEqual([liste, [], []]);
    expect(horsPermis(await clicsExtension(page), PERMIS_CCI)).toEqual([]);
    expect((await clicsExtension(page)).filter((c) => c === "bascule-cci")).toEqual([]);
  });

  test("Cci + Letsignit + importance élevée + texte de signature retiré, ensemble", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { delai: 200 } });
    const liste = adresses(4, "cci");
    await preparerEnvoi(page, { dest: liste, objet: "Objet", texte: "Bonjour à tous,\nVenez nombreux.", cci: true, paquet: 10, max: 1000, importance: true, sansSignature: true });
    expect(await envoyer(page)).toContain("Terminé : 4 destinataire(s) servi(s) en 1 mail(s).");
    const [m] = await envoyes(page);
    expect([m!.cci, m!.a, m!.importance]).toEqual([liste, [], "haute"]);
    await verifierCorps(page, m!, "<div>Bonjour à tous,</div><div>Venez nombreux.</div>", "texte-retire");
    expect(horsPermis(await clicsExtension(page), PERMIS_CCI, PERMIS_IMPORTANCE)).toEqual([]);
  });
});

// ================================================================== IMPORTANCE
test.describe("importance élevée (E4)", () => {
  test("bouton « Importance élevée » du ruban (aria-pressed) : confirmé, envoyé en importance haute", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr", "jean.dupont@ecole-test.fr"], objet: "Objet", texte: "Texte", importance: true });
    expect(await envoyer(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    expect((await envoyes(page)).map((m) => m.importance)).toEqual(["haute", "haute"]);
    expect(horsPermis(await clicsExtension(page), PERMIS_A, PERMIS_IMPORTANCE)).toEqual([]);
  });

  test("ruban resté sur « Accueil » à l'ouverture de la rédaction (variante) : onglet « Message » puis « Importance élevée »", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { ongletRedaction: "Accueil" });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte", importance: true });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    expect((await envoyes(page))[0]!.importance).toBe("haute");
    expect(horsPermis(await clicsExtension(page), PERMIS_A, PERMIS_IMPORTANCE)).toEqual([]);
  });

  test("aria-pressed ne suit pas (Outlook n'affiche pas l'état) : rien n'est envoyé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { importanceSansEtat: true });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte", importance: true });
    expect(await envoyer(page)).toContain("rien n'est envoyé");
    expect(await envoyes(page)).toEqual([]);
  });

  test("fenêtre étroite : « Importance élevée » rangée dans « Autres options » du ruban : importance confirmée par le menu, ou arrêt sans envoi ; jamais d'autre commande", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { rubanEtroit: true });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], objet: "Objet", texte: "Texte", importance: true });
    const r = await envoyer(page);
    const m = await envoyes(page);
    if (m.length) expect(m[0]!.importance).toBe("haute");
    else expect(r).toContain("rien n'est envoyé");
    const autres = horsPermis(await clicsExtension(page), PERMIS_A, PERMIS_IMPORTANCE, ["ruban:autres-options", "debord:importance-haute"]);
    // BUG : importanceHaute (publipostage/outlook.js:693-703) passe à l'onglet « Options » AVANT
    // d'ouvrir « Autres options » : il ouvre alors le débordement de l'onglet Options (sans
    // importance), jamais celui de l'onglet Message où la commande est rangée.
    expect({ envoyes: m.length, autres }, r).toEqual({ envoyes: 1, autres: [] });
  });
});

// ================================================================== SIGNATURE
test.describe("retrait du texte de la signature Letsignit", () => {
  test("signature présente avant le collage : texte retiré, logo, réseaux et bannière gardés, message intact", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { delai: 0 } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr", "jean.dupont@ecole-test.fr"], objet: "Objet", texte: "Chers étudiants,\nRendez-vous au Campus de Paris le 27.", sansSignature: true });
    expect(await envoyer(page)).toContain("Terminé : 2 mail(s) envoyé(s).");
    for (const m of await envoyes(page)) {
      await verifierCorps(page, m, "<div>Chers étudiants,</div><div>Rendez-vous au Campus de Paris le 27.</div>", "texte-retire");
      expect(m.html).toContain('href="https://www.linkedin.com/school/ecole-test"');
    }
  });

  test("signature posée APRÈS le collage (dans #Signature) : retirée avant l'envoi", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { declencheur: "collage-corps", delai: 0 } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], sansSignature: true });
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s).");
    const modele = await pp(page).evaluate((h) => (h as HTMLElement).shadowRoot!.getElementById("pp-texte")!.innerHTML);
    await verifierCorps(page, (await envoyes(page))[0]!, modele, "texte-retire");
  });

  test("variante plausible : signature posée après le collage SANS bloc #Signature : jamais d'envoi avec le texte de la signature (E3)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { letsignit: { declencheur: "collage-corps", delai: 0, conteneur: null } });
    await preparerEnvoi(page, { dest: ["alice.martin@ecole-test.fr"], sansSignature: true });
    const r = await envoyer(page);
    const m = await envoyes(page);
    // BUG : retirerTexteSignature (publipostage/outlook.js:507-535) ne connaît que le texte présent
    // AVANT le collage et les blocs [id*=signature] : une signature ajoutée ensuite sans ce bloc
    // n'est ni retirée ni signalée (contrôle final « ok »), le mail part avec nom, poste, ville.
    for (const x of m) for (const t of SIGNATURE_TEXTES) expect(x.texte, "envoyé avec « " + t + " » : " + r).not.toContain(t);
  });
});

// ================================================================== PAUSE, RECHARGEMENT, RYTHME
test.describe("pause, rechargement, rythme (E2, E7)", () => {
  test("pause pendant la pause entre deux mails, rechargement, reprise : chacun une fois, dans l'ordre", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["p1@ecole-test.fr", "p2@ecole-test.fr", "p3@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 600, pmax: 600 });
    const avant = (await suivi(page, "__ppSuivi")).fins;
    await lancerSansAttendre(page);
    await expect.poll(async () => (await envoyes(page)).length, { timeout: 20000 }).toBe(1);
    await pp(page).locator("#pp-pause").click();
    expect(await finEnvoi(page, avant)).toContain("Envoi en pause");
    const premier = await envoyes(page);
    await recharger(page);
    await pp(page).locator("#pp-pmin").fill("0");
    await pp(page).locator("#pp-pmax").fill("0");
    expect(await envoyer(page)).toContain("Terminé : 3 mail(s) envoyé(s)."); // bilan de toute la campagne
    expect(premier.concat(await envoyes(page)).map((m) => m.a[0])).toEqual(["p1@ecole-test.fr", "p2@ecole-test.fr", "p3@ecole-test.fr"]);
  });

  test("rechargement pendant « Envoyer » (la rédaction n'est pas encore refermée) : « à vérifier », jamais renvoyé à la reprise", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { envoiLent: 100000 }); // 1 s en mode accéléré : la rédaction reste ouverte
    await preparerEnvoi(page, { dest: ["r1@ecole-test.fr", "r2@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    await lancerSansAttendre(page);
    await page.waitForFunction(() => (window as unknown as { __journal: Journal }).__journal.some((e) => e.type === "envoyer-clique"), null, { polling: 5, timeout: 20000 });
    await recharger(page); // le mail est peut-être parti, peut-être pas
    await page.evaluate(() => { (window as unknown as { __cfg: { envoiLent: number } }).__cfg.envoiLent = 0; });
    await expect(pp(page).locator("#pp-suivi-resume")).toHaveText("0 envoyé(s) sur 2, 1 à vérifier");
    expect(await envoyer(page)).toContain("Terminé : 1 mail(s) envoyé(s), 1 à vérifier.");
    expect((await envoyes(page)).map((m) => m.a[0])).toEqual(["r2@ecole-test.fr"]);
  });

  test("limite de 2 mails sur 24 h tenue après rechargement", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["l1@ecole-test.fr", "l2@ecole-test.fr", "l3@ecole-test.fr"], objet: "Objet", texte: "Texte", max: 2 });
    expect(await envoyer(page)).toContain("Limite de 2 mails sur 24 h atteinte");
    await recharger(page);
    expect(await envoyer(page)).toContain("Limite de 2 mails sur 24 h atteinte");
    expect((await envoyes(page)).length).toBe(0); // nouvelle page : la fausse page repart de zéro
  });

  test("arrière-plan (fenêtre sans focus, minuteries de l'extension bridées à une par minute) : l'envoi avance, pauses tenues, rien de travers", async ({ context, worker }) => {
    // Chromium headless garde visibilityState « visible » : on simule l'arrière-plan par la perte
    // du focus de la fenêtre et le bridage des minuteries du monde isolé de l'extension (comme
    // publipostage.spec.ts), ce que fait Chrome pour un onglet caché.
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["f1@ecole-test.fr", "f2@ecole-test.fr", "f3@ecole-test.fr"], objet: "Arrière-plan", texte: "Texte", pmin: 100, pmax: 100 }); // 1 s
    const cdp = await context.newCDPSession(page);
    const mondes: Array<{ id: number; name: string; auxData?: { type?: string } }> = [];
    cdp.on("Runtime.executionContextCreated", (e) => mondes.push(e.context));
    await cdp.send("Runtime.enable");
    const monde = mondes.find((c) => c.auxData?.type === "isolated" && c.name === "Smart Mail");
    expect(monde).toBeTruthy();
    await cdp.send("Runtime.evaluate", { contextId: monde!.id, expression: "(() => { const st = window.setTimeout; window.setTimeout = (f, ms, ...a) => st(f, (ms || 0) + 60000, ...a); })()" });
    const autre = await context.newPage();
    await autre.goto("about:blank");
    await autre.bringToFront();
    await cdp.send("Emulation.setFocusEmulationEnabled", { enabled: false });
    expect(await page.evaluate(() => document.hasFocus())).toBe(false);
    const avant = (await suivi(page, "__ppSuivi")).fins;
    const t0 = Date.now();
    await lancerSansAttendre(page);
    const r = await finEnvoi(page, avant, 90000);
    expect(r).toContain("Terminé : 3 mail(s) envoyé(s).");
    expect(Date.now() - t0).toBeLessThan(60000);
    const m = await envoyes(page);
    expect(m.map((x) => x.a[0])).toEqual(["f1@ecole-test.fr", "f2@ecole-test.fr", "f3@ecole-test.fr"]);
    for (const x of m) { expect([x.cc, x.cci, x.libres, x.objet]).toEqual([[], [], [], "Arrière-plan"]); expect(compact(x.texte)).toBe("Texte"); }
    for (let i = 1; i < m.length; i++) expect(m[i]!.t - m[i - 1]!.t).toBeGreaterThanOrEqual(1000);
    expect(horsPermis(await clicsExtension(page), PERMIS_A)).toEqual([]);
  });
});

// ================================================================== DIAGNOSTIC
test.describe("diagnostic de l'onglet Envoyer", () => {
  test("le diagnostic ne tape rien, n'envoie rien, n'ouvre pas le carnet, et referme son brouillon (bouton « Abandonner »)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await pp(page).locator("#pp-gear").click();
    const [dl] = await Promise.all([page.waitForEvent("download"), pp(page).locator("#pp-diag").click()]);
    const rapport = JSON.parse(await (await import("node:fs/promises")).readFile((await dl.path())!, "utf8"));
    expect(await envoyes(page)).toEqual([]);
    expect((await etat(page)).carnet).toBe(0);
    expect(ecritures(await journal(page))).toEqual([]);
    // BUG : SEL.ignorer (publipostage/outlook.js:46) ne connaît que « Ignorer »/« Discard » ;
    // le vrai Outlook dit « Abandonner » : le diagnostic laisse son brouillon ouvert (importance
    // haute déjà cochée), et le prochain envoi s'arrête « Un message est déjà ouvert ».
    expect({ refermee: rapport.refermee, composes: (await etat(page)).composes.length }).toEqual({ refermee: true, composes: 0 });
  });
});

// ================================================================== RELEVÉ
async function attendreRecherches(page: Page, n: number, timeout = 60000) {
  await expect.poll(() => page.evaluate(() => (window as unknown as { __typed: string[] }).__typed.length), { timeout }).toBeGreaterThanOrEqual(n);
}
const touchesExtension = (j: Journal, depuis = 0) => j.filter((e) => e.n >= depuis && e.type === "touche" && !e.confiance && e.cible !== "PANNEAU").map((e) => e.touche!);

test.describe("relevé d'adresses (R1–R6)", () => {
  test("« Démarrer » depuis la boîte de réception : nouveau message, frappe dans « À » seulement, rien de validé ni d'envoyé, « À » vide à la fin, adresses exactes", async ({ context, worker }) => {
    test.setTimeout(300000);
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    const msg = await finReleve(page, d, 280000);
    expect(msg).toContain("Terminé");
    const j = await journal(page);
    const base = await baseReleve(worker);
    expect(base.sort()).toEqual(ANNUAIRE.map((x) => x[1]).sort()); // R3 : tout l'annuaire, rien d'autre
    for (const a of ADRESSES_HORS_ANNUAIRE) expect(base).not.toContain(a);
    expect(pastillesCreees(j)).toEqual([]); // R1
    expect(touchesExtension(j).filter((k) => !/^([a-z0-9 ]|Backspace)$/.test(k))).toEqual([]); // R1
    expect(new Set(ecritures(j))).toEqual(new Set(["A"])); // R4
    expect(horsPermis(await clicsExtension(page), ["nouveau", "A", "plus-de-resultats"])).toEqual([]); // R1, E6
    const [c] = (await etat(page)).composes;
    expect([c!.a.length, c!.libres.A]).toEqual([0, ""]); // R2
    expect(await envoyes(page)).toEqual([]);
    // Noms bien associés (Fluent v9 : spans collés, badge « Externe », adresse répétée).
    const lignes = await ao(page).locator("#ao-rows .res").evaluateAll((rs) => rs.map((r) => [r.querySelector(".m")?.textContent, r.querySelector(".n")?.textContent ?? ""]));
    const noms = new Map(lignes.map(([m, n]) => [m, n]));
    expect(noms.get("abder@example.com")).toBe("Miriam ABDER");
    expect(noms.get("alice.prunier@partenaire.example")).toBe("Alice PRUNIER");
    expect(noms.get("lea.durand@ecole-test.fr")).toBe("DURAND Léa");
  });

  test("« Démarrer » alors que le curseur est dans le CORPS de la réponse que l'utilisateur écrit : jamais de frappe dans le corps, texte intact", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { reponse: { depuis: 1 } }, { onglet: "adresses" });
    await page.locator('[data-ident="CORPS"]').click({ position: { x: 30, y: 8 } }); // Répondre met le curseur dans le corps
    await page.keyboard.type("Merci Alice, je serai là lundi.");
    const avant = (await etat(page)).composes[0]!;
    const n0 = (await journal(page)).length;
    await demarrerReleve(page);
    await page.waitForTimeout(6000);
    await ao(page).locator("#ao-stop").click().catch(() => {});
    await page.waitForTimeout(1500);
    const apres = (await etat(page)).composes[0]!;
    // BUG : champActif (adresses/panel.js:465-472) n'écarte que recherche/objet : le corps
    // (« Corps du message ») est pris pour le champ du relevé (armHere, adresses/armement.js:132-145),
    // puis le 1er effacement (adresses/adapter-dom.js:99-112, 200 retours arrière) mange le texte
    // de l'utilisateur et les préfixes sont tapés dans le corps.
    const horsDeA = ecritures(await journal(page), n0).filter((c) => c !== "A");
    expect({
      debutDuCorps: compact(apres.corps).slice(0, 40),
      frappesHorsDeA: horsDeA.length,
    }).toEqual({ debutDuCorps: compact(avant.corps).slice(0, 40), frappesHorsDeA: 0 });
    expect(compact(apres.corps)).toBe(compact(avant.corps));
  });

  test("« Démarrer » avec le curseur dans « À » d'un brouillon qui a déjà des destinataires : les destinataires de l'utilisateur restent", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { brouillon: { a: ["paul.robert@ecole-test.fr", "emma.richard@ecole-test.fr"], objet: "Projet", corps: "Bonjour à vous deux." } }, { onglet: "adresses" });
    await page.locator('[data-ident="A"]').click({ position: { x: 400, y: 8 } }); // clic à droite des pastilles : curseur à la fin
    await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    await ao(page).locator("#ao-stop").click();
    await page.waitForTimeout(1500);
    const [c] = (await etat(page)).composes;
    // BUG : le 1er effacement du passage (adresses/adapter-dom.js:103, n porté à 200) efface tout
    // le champ jusqu'à le vider, pastilles comprises (execCommand("delete") retire une pastille
    // contenteditable=false d'un coup) : les destinataires que l'utilisateur avait mis disparaissent.
    expect(c!.a.map((x) => x.adresse)).toEqual(["paul.robert@ecole-test.fr", "emma.richard@ecole-test.fr"]);
    expect(c!.libres.A).toBe("");
  });

  test("« Démarrer » sans curseur avec une réponse ouverte (À pré-rempli) : la pastille de l'expéditeur reste, rien ne reste tapé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { reponse: { depuis: 1 } }, { onglet: "adresses" });
    await page.locator('[data-ident="dossier:reception"]').click(); // le focus quitte la réponse (rien d'éditable)
    await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    await ao(page).locator("#ao-stop").click();
    await page.waitForTimeout(1500);
    const [c] = (await etat(page)).composes;
    expect(c!.a.map((x) => x.adresse)).toEqual(["alice.martin@ecole-test.fr"]);
    expect(c!.libres.A).toBe("");
    expect(ecritures(await journal(page)).filter((x) => x !== "A")).toEqual([]);
  });

  test("pendant le relevé, l'utilisateur clique dans la recherche et tape : arrêt propre, le focus n'est pas repris, sa recherche intacte (R4)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    const n0 = (await journal(page)).length;
    await page.locator('[data-ident="RECHERCHE"]').click();
    await page.keyboard.type("rapport annuel", { delay: 120 });
    await page.waitForTimeout(1500);
    const valeur = await page.locator('[data-ident="RECHERCHE"]').inputValue();
    const actif = (await etat(page)).actif;
    const enCours = !(await ao(page).locator("#ao-compact-view").isHidden());
    if (enCours) await ao(page).locator("#ao-stop").click();
    await finReleve(page, d, 20000);
    const j = await journal(page);
    // Touches de l'UTILISATEUR (événements « isTrusted ») arrivées dans « À » au lieu de la recherche.
    const volees = j.filter((e) => e.n >= n0 && e.type === "touche" && e.confiance && e.cible === "A").map((e) => e.touche).join("");
    const [c] = (await etat(page)).composes;
    // BUG : focus() (adresses/adapter-dom.js:130-137) refait el.focus() à chaque recherche, et
    // taper()/insertText (adapter-dom.js:73-97) aussi à chaque lettre : le relevé reprend le focus
    // au lieu de s'arrêter, et la saisie de l'utilisateur part en partie dans « À ».
    expect({ valeur, actif, enCours, volees }).toEqual({ valeur: "rapport annuel", actif: "RECHERCHE", enCours: false, volees: "" });
    expect([c!.a.length, c!.libres.A]).toEqual([0, ""]);
  });

  test("une pastille apparaît dans « À » pendant le relevé (Entrée tapée alors que le curseur y est) : le relevé s'arrête aussitôt et ne l'efface pas (filet R1)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    // Un préfixe est tapé et la liste propose des contacts : Entrée valide la suggestion surlignée.
    await page.evaluate(() => new Promise<void>((ok) => {
      const essai = () => {
        const a = document.querySelector('[data-ident="A"]') as HTMLElement | null;
        const pop = document.getElementById("suggestions-destinataires") as HTMLElement | null;
        const libre = a ? (a.textContent ?? "").replace(/[\u200b]/g, "").trim() : "";
        if (a && libre && pop && !pop.hidden && pop.querySelector('[role="option"][aria-selected="true"]') && document.activeElement === a) {
          a.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true, cancelable: true }));
          ok();
        } else setTimeout(essai, 2);
      };
      essai();
    }));
    const valide = (await journal(page)).filter((e) => e.type === "pastille" && e.source === "touche Enter");
    expect(valide.length, "la fausse page a bien validé une suggestion").toBe(1);
    await page.waitForTimeout(5000);
    const enCours = !(await ao(page).locator("#ao-compact-view").isHidden());
    if (enCours) await ao(page).locator("#ao-stop").click();
    const msg = await finReleve(page, d, 20000);
    const [c] = (await etat(page)).composes;
    // BUG : chipSig/newChipAppeared (adresses/adapter-dom.js:308-323) ne cherche que
    // [data-hovercard-id], [role=listitem], .ms-PickerPersona-container, [data-selection-index] :
    // la vraie pastille du nouvel Outlook (span._EType_RECIPIENT_ENTITY, id REK…) est invisible
    // pour le filet ; le relevé continue et l'efface à coups de retour arrière (effacer).
    expect({ enCours, pastilles: c!.a.map((x) => x.adresse) }, msg).toEqual({ enCours: false, pastilles: [valide[0]!.adresse] });
  });

  test("brouillon de l'utilisateur SANS destinataire (objet et texte écrits), curseur dans « À » : seul « À » sert, objet et texte intacts, « À » vide à la fin", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { brouillon: { a: [], objet: "Compte rendu", corps: "Voici le compte rendu de la réunion.", focus: "A" } }, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 3, 30000);
    await ao(page).locator("#ao-stop").click();
    expect(await finReleve(page, d, 20000)).toContain("En pause");
    const [c] = (await etat(page)).composes;
    expect([c!.objet, compact(c!.corps), c!.a.length, c!.libres.A]).toEqual(["Compte rendu", compact("Voici le compte rendu de la réunion."), 0, ""]);
    expect(new Set(ecritures(await journal(page)))).toEqual(new Set(["A"]));
  });

  test("« Démarrer » avec le curseur dans « Cc » d'un nouveau message : tout se passe dans Cc, « À » intact, Cc vidé à la fin", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    await page.evaluate(() => (window as unknown as { __fausse: { nouveau: (f: string) => boolean } }).__fausse.nouveau("A"));
    await page.locator('[data-ident="bascule-cc"]').dispatchEvent("click");
    await page.locator('[data-ident="CC"]').click({ position: { x: 20, y: 8 } });
    const n0 = (await journal(page)).length;
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 0, 1000).catch(() => {});
    await expect.poll(async () => ecritures(await journal(page), n0).length, { timeout: 30000 }).toBeGreaterThan(4);
    await ao(page).locator("#ao-stop").click();
    await finReleve(page, d, 20000);
    const [c] = (await etat(page)).composes;
    expect(new Set(ecritures(await journal(page), n0))).toEqual(new Set(["CC"]));
    expect([c!.a.length, c!.libres.A, c!.cc.length, c!.libres.CC]).toEqual([0, "", 0, ""]);
  });

  test("« Démarrer » avec le curseur dans l'objet d'un nouveau message : l'objet n'est jamais touché, le relevé se fait dans « À »", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    await page.evaluate(() => (window as unknown as { __fausse: { nouveau: (f: string) => boolean } }).__fausse.nouveau("OBJET"));
    await page.locator('[data-ident="OBJET"]').fill("Sortie du jeudi");
    const n0 = (await journal(page)).length;
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    await ao(page).locator("#ao-stop").click();
    await finReleve(page, d, 20000);
    const [c] = (await etat(page)).composes;
    expect(new Set(ecritures(await journal(page), n0))).toEqual(new Set(["A"]));
    expect([c!.objet, c!.a.length, c!.libres.A]).toEqual(["Sortie du jeudi", 0, ""]);
  });

  test("« Démarrer » avec le curseur dans la barre de recherche : jamais de frappe dans la recherche, un nouveau message est ouvert", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    await page.locator('[data-ident="RECHERCHE"]').click();
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    await ao(page).locator("#ao-stop").click();
    await finReleve(page, d, 20000);
    expect(new Set(ecritures(await journal(page)))).toEqual(new Set(["A"]));
    expect(await page.locator('[data-ident="RECHERCHE"]').inputValue()).toBe("");
    expect((await etat(page)).composes.map((c) => c.type)).toEqual(["nouveau"]);
  });

  test("« Démarrer » alors que le carnet d'adresses « Ajouter des destinataires » est ouvert : rien n'est tapé, message clair", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    await page.evaluate(() => (window as unknown as { __fausse: { nouveau: (f: string) => boolean } }).__fausse.nouveau("A"));
    await page.locator('[data-ident="libelle-a"]').click(); // le libellé « À » ouvre le carnet
    await expect(page.locator('[data-ident="dialogue:carnet"]')).toBeVisible();
    const n0 = (await journal(page)).length;
    const d = await demarrerReleve(page);
    const msg = await finReleve(page, d, 30000);
    expect(ecritures(await journal(page), n0)).toEqual([]);
    expect(msg).toContain("Rien n'a été tapé");
    expect((await etat(page)).modales).toEqual(["carnet"]); // rien cliqué dans le carnet
  });

  test("liste qui ne montre d'abord AUCUN contact, seulement « Afficher plus de résultats » (comme owa-reel) : adresses et noms exacts, aucun clic sur un contact", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, { initial: 0, lot: 2, suggestionsAuFocus: false }, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 6, 60000);
    await ao(page).locator("#ao-stop").click();
    await finReleve(page, d, 20000);
    const base = await baseReleve(worker);
    expect(base.length).toBeGreaterThan(3);
    for (const a of base) expect(ANNUAIRE.map((x) => x[1])).toContain(a);
    const j = await journal(page);
    expect(horsPermis(await clicsExtension(page), ["plus-de-resultats", "nouveau"])).toEqual([]);
    expect(j.filter((e) => e.type === "plus-de-resultats").length).toBeGreaterThan(0);
    expect(pastillesCreees(j)).toEqual([]);
    const lignes = await ao(page).locator("#ao-rows .res").evaluateAll((rs) => rs.map((r) => [r.querySelector(".m")?.textContent ?? "", r.querySelector(".n")?.textContent ?? ""]));
    for (const [mail, nom] of lignes) expect(nom, mail).toBe(ANNUAIRE.find((x) => x[1] === mail)![0]);
  });

  for (const flou of ["pastille", "texte"] as const) {
    test(`fenêtre « Letsignit Add-in » (piège à focus Fluent) pendant le relevé, Outlook laissant le texte quitté en ${flou} : « À » laissé vide (R1, R2)`, async ({ context, worker }) => {
      const page = await ouvrir(context, worker, { flou }, { onglet: "adresses" });
      const d = await demarrerReleve(page);
      await attendreRecherches(page, 2, 30000);
      // Le complément s'affiche pendant la rédaction (comme sur le vrai Outlook), 3 s, pendant qu'un préfixe est tapé.
      await page.waitForFunction(() => { const a = document.querySelector('[data-ident="A"]'); return !!a && (a.textContent ?? "").replace(/[\u200b]/g, "").trim().length > 0; }, null, { polling: 5 });
      await page.evaluate(() => (window as unknown as { __fausse: { modale: (o: unknown) => void } }).__fausse.modale({ id: "letsignit", titre: "Letsignit Add-in", texte: "Letsignit : mise à jour de votre signature…", piege: true }));
      await page.waitForTimeout(3000);
      await page.evaluate(() => (window as unknown as { __fausse: { fermerModales: () => void } }).__fausse.fermerModales());
      await page.waitForTimeout(2500);
      if (!(await ao(page).locator("#ao-compact-view").isHidden())) await ao(page).locator("#ao-stop").click();
      const msg = await finReleve(page, d, 30000);
      const [c] = (await etat(page)).composes;
      // BUG : arrêt « le curseur a quitté le champ « À » » (adresses/core/moteur.js:504) sans
      // nettoyage : le nettoyage final n'a lieu que si focus() réussit (moteur.js:542), et rien
      // n'attend les fenêtres d'Outlook côté relevé (pas d'équivalent de attendreFenetre) :
      // le préfixe tapé reste dans « À » (en pastille si Outlook valide au départ du focus).
      expect({ pastilles: c!.a.map((x) => x.texte ?? x.adresse), libre: c!.libres.A }, msg).toEqual({ pastilles: [], libre: "" });
    });
  }

  test("variante plausible : une fenêtre d'Outlook avec un champ (enquête) s'ouvre juste avant la frappe d'un préfixe : jamais de frappe dans la fenêtre (R4)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    // Dès que le relevé a effacé sa recherche (champ vide, juste avant de taper la suivante).
    await page.evaluate(() => {
      const a = document.querySelector('[data-ident="A"]') as HTMLElement;
      let fait = false;
      a.addEventListener("input", () => {
        if (fait || (a.textContent ?? "").replace(/[​]/g, "").trim()) return;
        fait = true;
        (window as unknown as { __fausse: { modale: (o: unknown) => void } }).__fausse.modale({ id: "enquete", titre: "Votre avis sur Outlook", texte: "Quelle est la probabilité que vous recommandiez Outlook ?", champ: "Commentaires", piege: true, boutons: [{ texte: "Plus tard", ident: "dialogue:enquete:plus-tard" }] });
        setTimeout(() => (window as unknown as { __fausse: { fermerModales: () => void } }).__fausse.fermerModales(), 3000);
      });
    });
    await expect.poll(async () => (await journal(page)).some((e) => e.type === "modale" && e.id === "enquete"), { timeout: 30000 }).toBe(true);
    await page.waitForTimeout(4500);
    if (!(await ao(page).locator("#ao-compact-view").isHidden())) await ao(page).locator("#ao-stop").click();
    const msg = await finReleve(page, d, 30000);
    const dansLaFenetre = (await journal(page)).filter((e) => (e.type === "saisie" || e.type === "collage") && String(e.cible).startsWith("dialogue:"));
    // BUG : taper → insertText (adresses/adapter-dom.js:73-76) fait el.focus() puis
    // execCommand("insertText") sans vérifier que le focus est bien arrivé dans « À » : la
    // fenêtre l'a gardé (piège à focus), la lettre du relevé est tapée dans son champ.
    expect(dansLaFenetre.map((e) => e.cible + " ← " + e.data), msg).toEqual([]);
  });

  test("pause pendant une recherche : « À » vidé, rien de validé (R2)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 3, 30000);
    await ao(page).locator("#ao-stop").click();
    expect(await finReleve(page, d, 20000)).toContain("En pause");
    const [c] = (await etat(page)).composes;
    expect([c!.a.length, c!.libres.A]).toEqual([0, ""]);
    expect(pastillesCreees(await journal(page))).toEqual([]);
  });

  test("l'utilisateur clique « Abandonner » pendant le relevé : arrêt propre, rien n'est tapé ailleurs ensuite (R4)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    await page.locator('[data-ident="abandonner"]').dispatchEvent("click"); // (le panneau couvre le bouton à l'écran)
    const msg = await finReleve(page, d, 20000);
    expect(msg).toMatch(/Arrêté|En pause/);
    const j = await journal(page);
    const abandon = j.find((e) => e.type === "abandon")!; // instant exact de la fermeture du message
    expect(ecritures(j, abandon.n)).toEqual([]);
    expect(await envoyes(page)).toEqual([]);
  });

  test("la fenêtre d'Outlook perd le focus pendant le relevé (autre application devant) : arrêt dit clairement, « À » vidé (R6)", async ({ context, worker }) => {
    const page = await ouvrir(context, worker, {}, { onglet: "adresses" });
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 2, 30000);
    const cdp = await context.newCDPSession(page);
    const autre = await context.newPage();
    await autre.goto("about:blank");
    await autre.bringToFront();
    await cdp.send("Emulation.setFocusEmulationEnabled", { enabled: false });
    expect(await page.evaluate(() => document.hasFocus())).toBe(false);
    const msg = await finReleve(page, d, 30000);
    expect(msg).toContain("n'est plus au premier plan");
    const [c] = (await etat(page)).composes;
    expect([c!.a.length, c!.libres.A]).toEqual([0, ""]);
    expect(pastillesCreees(await journal(page))).toEqual([]);
  });
});

// ================================================================== VERROU MUTUEL
test.describe("verrou mutuel relevé / envoi (M1)", () => {
  test("clics simultanés sur « Démarrer » et « Lancer l'envoi » : un seul tourne, la frappe du relevé n'atteint jamais le mail envoyé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["m1@ecole-test.fr", "m2@ecole-test.fr"], objet: "Objet", texte: "Texte" });
    await page.evaluate(() => {
      const r = (s: string) => (document.querySelector(s) as HTMLElement & { shadowRoot: ShadowRoot }).shadowRoot;
      (r('[data-testid="ao-panel"]').getElementById("ao-run") as HTMLButtonElement).click();
      (r("[data-pp-panneau]").getElementById("pp-go") as HTMLButtonElement).click();
    });
    await page.waitForTimeout(8000);
    const m = await envoyes(page);
    const releve = await page.evaluate(() => (window as unknown as { __typed: string[] }).__typed.length);
    for (const x of m) { expect(x.a.length + x.cci.length).toBe(1); expect(x.libres).toEqual([]); }
    expect(m.length === 0 || releve === 0, `envoi (${m.length}) ET relevé (${releve} recherches) ont tourné ensemble`).toBe(true);
    if (!(await ao(page).locator("#ao-compact-view").isHidden())) await ao(page).locator("#ao-stop").click().catch(() => {});
  });

  test("pendant un envoi, « Démarrer » est refusé ; pendant un relevé, « Lancer l'envoi » est refusé", async ({ context, worker }) => {
    const page = await ouvrir(context, worker);
    await preparerEnvoi(page, { dest: ["v1@ecole-test.fr", "v2@ecole-test.fr"], objet: "Objet", texte: "Texte", pmin: 1000, pmax: 1000 });
    await lancerSansAttendre(page);
    await expect.poll(async () => (await envoyes(page)).length, { timeout: 20000 }).toBe(1);
    await onglet(page, "pp", "adresses").click();
    await ao(page).locator("#ao-run").click();
    await expect(ao(page).locator("#ao-notice")).toContainText("Un envoi Publipostage est en cours");
    expect(await page.evaluate(() => (window as unknown as { __typed: string[] }).__typed.length)).toBe(0);
    await onglet(page, "ao", "envoi").click();
    await pp(page).locator("#pp-pause").click();
    await expect(pp(page).locator("#pp-notice")).toContainText("Envoi en pause", { timeout: 20000 });
    await onglet(page, "pp", "adresses").click();
    const d = await demarrerReleve(page);
    await attendreRecherches(page, 1, 30000);
    await onglet(page, "ao", "envoi").click();
    await pp(page).locator("#pp-go").click();
    await expect(pp(page).locator("#pp-notice")).toContainText("Un relevé d'adresses est en cours");
    await onglet(page, "pp", "adresses").click();
    await ao(page).locator("#ao-stop").click();
    await finReleve(page, d, 20000);
    expect((await envoyes(page)).length).toBe(1);
  });
});
