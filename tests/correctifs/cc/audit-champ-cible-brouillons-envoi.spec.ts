// Audit Smart Mail — « Champ visé, brouillons de l'utilisateur et vol du focus » (3/3) : l'ENVOI.
// C : lancé depuis chaque état d'Outlook. D : pendant un mail, l'utilisateur ou Outlook change la rédaction visible,
// la ferme, change de vue, re-rend « À », vole le focus. Fausse page en mode accéléré (attentes /100) ; les
// déclencheurs « ow.piege » rejouent de façon déterministe une action de l'utilisateur à un instant précis du mail.
// Règles : E1, E3, E5, E6 (+ R5 : jamais modifier le brouillon de l'utilisateur).
import { test, expect, pp, ouvrir, etat, clicsExtension, saisiesExtension, lancerEnvoi, attendreFinEnvoi, URL_OUTLOOK,
  contenuUtilisateur } from "./audit/champ-cible-brouillons/banc.ts";
import type { Page } from "@playwright/test";

test.describe.configure({ mode: "parallel" });

const DEST = "alice.martin@exemple.fr\nbruno.faure@exemple.fr";
const suivi = async (page: Page) => ((await pp(page).locator("#pp-suivi-resume").textContent()) || "").trim();
const nbEnvoyeSelonSuivi = async (page: Page) => parseInt(((await suivi(page)).match(/(\d+) envoyé/) || [])[1] || "0", 10);
// Les mails partis depuis une rédaction de l'EXTENSION : bons destinataires, objet et corps exacts.
function verifierMailsExtension(e: Awaited<ReturnType<typeof etat>>, attendus: string[]) {
  const m = e.envoyes.filter((x) => x.proprietaire === "extension");
  expect(m.map((x) => x.a)).toEqual(attendus.map((a) => [a]));
  for (const x of m) {
    expect(x.cc).toEqual([]);
    expect(x.objet).toBe("Invitation");
    expect(x.corps.split("Cordialement")[0].replace(/\s+/g, " ").trim()).toBe("Bonjour, Voici l'invitation. À bientôt");
  }
}
async function brouillonUtilisateur(page: Page, o: { pastilles?: string[]; objet?: string; corps?: string } = {}) {
  return page.evaluate((o) => {
    const ow = (window as any).ow; const id = ow.nouvelleRedaction("utilisateur");
    (o.pastilles ?? ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]).forEach((a: string) => ow.ajouterPastille(id, "À", a));
    if (o.objet !== "") ow.ecrireObjet(id, o.objet ?? "Budget 2027");
    ow.ecrireCorps(id, o.corps ?? "Bonjour à tous, vous trouverez ci-joint le budget prévisionnel."); return id as string;
  }, o);
}

test.describe("Envoi — lancé depuis les états d'Outlook", () => {
  test("C01 boîte de réception : un mail par destinataire, rien d'autre touché", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    verifierMailsExtension(await etat(page), ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]);
  });

  test("C02 message lu affiché : envoi normal, le message lu n'est jamais « répondu »", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(1));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    const e = await etat(page);
    verifierMailsExtension(e, ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]);
    expect((await clicsExtension(page)).map((c) => c.libelle)).not.toContain("Répondre");
  });

  test("C03 réponse en ligne de l'utilisateur ouverte : refus clair, rien tapé ni cliqué dedans, rien envoyé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    const id = await page.evaluate(() => { const ow = (window as any).ow; const id = ow.repondre(0, "utilisateur"); ow.ecrireCorps(id, "Merci Paul, je regarde."); return id; });
    const avant = contenuUtilisateur(await etat(page));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/déjà ouvert en rédaction/);
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(contenuUtilisateur(e)).toEqual(avant);
    expect((await saisiesExtension(page)).map((s) => s.zone)).toEqual([]);
    expect((await clicsExtension(page)).filter((c) => c.zone !== "?" )).toEqual([]);
    void id;
  });

  test("C04 brouillon de l'utilisateur ouvert (À, objet, corps) : refus clair, brouillon intact", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await brouillonUtilisateur(page);
    const avant = contenuUtilisateur(await etat(page));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/déjà ouvert en rédaction/);
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(contenuUtilisateur(e)).toEqual(avant);
  });

  test("C05 deux brouillons en onglets, l'actif est celui de l'utilisateur : refus clair", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await brouillonUtilisateur(page, { objet: "Premier" });
    await brouillonUtilisateur(page, { objet: "Second", pastilles: ["celine.moreau@exemple.fr"] });
    const avant = contenuUtilisateur(await etat(page));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/déjà ouvert en rédaction/);
    expect(contenuUtilisateur(await etat(page))).toEqual(avant);
    expect((await etat(page)).envoyes).toEqual([]);
  });

  test("C06 brouillon de l'utilisateur dans un onglet masqué : l'envoi utilise SA rédaction, le brouillon reste intact et n'est jamais envoyé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await brouillonUtilisateur(page);
    await page.evaluate(() => (window as any).ow.afficherMessage(2));
    const avant = contenuUtilisateur(await etat(page));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    const e = await etat(page);
    verifierMailsExtension(e, ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]);
    expect(e.envoyes.filter((m) => m.proprietaire === "utilisateur")).toEqual([]);
    expect(contenuUtilisateur(e).redactions).toEqual(avant.redactions.map((r) => ({ ...r })));
  });

  test("C07 Calendrier : refus clair, rien tapé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.locator('nav#rail [data-vue="calendrier"]').click();
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/Nouveau message.*introuvable/);
    expect((await etat(page)).envoyes).toEqual([]);
    expect(await saisiesExtension(page)).toEqual([]);
  });

  test("C08 Contacts avec un formulaire en cours : refus clair, formulaire intact", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.locator('nav#rail [data-vue="contacts"]').click(); await page.locator("#nouveau-contact").click();
    await page.locator("#ct-prenom").fill("Jean");
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/introuvable/);
    expect((await etat(page)).contact.prenom).toBe("Jean");
  });

  test("C09 Paramètres ouverts (fenêtre modale) : refus clair après attente, rien tapé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.locator("#btn-parametres").click();
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/fenêtre par-dessus/);
    expect((await etat(page)).envoyes).toEqual([]);
    expect(await saisiesExtension(page)).toEqual([]);
  });

  test("C10 résultats de recherche affichés : envoi normal, recherche intacte", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.locator("#recherche").click(); await page.keyboard.type("planning"); await page.keyboard.press("Enter");
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    verifierMailsExtension(await etat(page), ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]);
    expect((await etat(page)).recherche).toBe("planning");
  });

  test("C11 discussion Teams ouverte avec un brouillon : envoi normal, Teams intact et jamais envoyé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.locator("#btn-teams").click(); await page.locator("#teams-saisie").click(); await page.keyboard.type("Je te rappelle");
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    const e = await etat(page);
    expect(e.teams).toBe("Je te rappelle");
    expect(e.teamsEnvoyes).toEqual([]);
  });

  test("C12 notification avec réponse rapide en cours : envoi normal, notification intacte", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.notifier("saisie")); await page.keyboard.type("ok");
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    expect((await etat(page)).toast).toBe("ok");
  });

  test("C13 après un envoi, Outlook ré-affiche le brouillon de l'utilisateur (onglet suivant) : arrêt clair, rien tapé dedans", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.apresFermeture = "onglet-suivant"; });
    await brouillonUtilisateur(page);
    await page.evaluate(() => (window as any).ow.afficherMessage(2));
    const avant = contenuUtilisateur(await etat(page));
    await lancerEnvoi(page, { dest: DEST });
    const n = await attendreFinEnvoi(page);
    const e = await etat(page);
    verifierMailsExtension(e, ["alice.martin@exemple.fr"]);
    // Arrêt prudent : soit « déjà ouvert en rédaction », soit « à vérifier » (la rédaction visible est celle de l'utilisateur).
    expect(n).toMatch(/déjà ouvert en rédaction|à vérifier|Vérifiez/);
    expect(e.envoyes.filter((m) => m.proprietaire === "utilisateur")).toEqual([]);
    expect(contenuUtilisateur(e)).toEqual(avant);
  });
});

test.describe("Envoi en cours — la rédaction visible change ou disparaît", () => {
  test("D01 l'utilisateur ouvre une réponse (et l'écrit) juste après le collage de l'objet : sa réponse n'est ni modifiée ni envoyée", async ({ page, worker }) => {
    // BUG : publipostage/outlook.js:146-150 champA/champObjet/champCorps re-cherchent à CHAQUE étape le premier champ
    // VISIBLE du document, sans garder la rédaction ouverte par ouvrirRedaction() : remplirCorps colle le message dans
    // la réponse de l'utilisateur, puis envoyer() (outlook.js:722-725, zoneRedaction() via champA()) clique SON « Envoyer ».
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "input", zone: /^objet-/, action: () => { const ow = (window as any).ow; const id = ow.repondre(0, "utilisateur"); ow.ecrireCorps(id, "Merci Paul, je regarde ça demain."); } }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes.filter((m) => m.proprietaire === "utilisateur").map((m) => ({ a: m.a, debut: m.corps.slice(0, 40) }))).toEqual([]);
    const rep = e.redactions.find((r) => r.proprietaire === "utilisateur");
    expect(rep?.corps ?? "(réponse partie)").toContain("Merci Paul, je regarde ça demain.");
    expect(rep?.corps ?? "").not.toContain("Voici l'invitation");
  });

  test("D02 … nouveau message de l'utilisateur (destinataire Claire, pas encore d'objet) ouvert juste après le collage dans « À » : jamais envoyé", async ({ page, worker }) => {
    // BUG : même cause que D01 ; en plus remplirA (outlook.js:367-372, aValidee) compte les pastilles du premier « À »
    // visible — celle de l'utilisateur — comme la validation de l'adresse collée : l'objet, le corps puis « Envoyer » vont
    // dans le message de l'utilisateur, envoyé à SON destinataire (E1, E3, E6), et le suivi note l'envoi à Alice.
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^a-/, action: () => { const ow = (window as any).ow; const id = ow.nouvelleRedaction("utilisateur"); ow.ajouterPastille(id, "À", "claire.petit@exemple.fr"); ow.ecrireCorps(id, "Claire, as-tu reçu le dossier ?"); } }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes.filter((m) => m.proprietaire === "utilisateur").map((m) => ({ a: m.a, objet: m.objet }))).toEqual([]);
    expect(e.envoyes.flatMap((m) => m.a)).not.toContain("claire.petit@exemple.fr");
  });

  test("D03 réponse ouverte juste après le collage du corps (avant « Envoyer ») : sa réponse n'est ni modifiée ni envoyée", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^corps-/, action: () => { const ow = (window as any).ow; const id = ow.repondre(0, "utilisateur"); ow.ecrireCorps(id, "Merci Paul, je regarde ça demain."); } }));
    await lancerEnvoi(page, { dest: DEST });
    const n = await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes.filter((m) => m.proprietaire === "utilisateur")).toEqual([]);
    const rep = e.redactions.find((r) => r.proprietaire === "utilisateur")!;
    expect(rep.corps).not.toContain("Voici l'invitation");
    expect(n).toMatch(/Arrêt avant l'envoi|rien n'est parti|rien n'est envoyé/);
  });

  test("D04 l'utilisateur clique un autre message juste après le collage de l'objet, discussion Teams ouverte : rien n'est collé ni envoyé dans Teams", async ({ page, worker }) => {
    // BUG : publipostage/outlook.js:150 champCorps() = premier champ visible de SEL.corps (outlook.js:42-43), dont
    // 'div[role="textbox"][contenteditable][aria-multiline]' : la rédaction masquée, c'est la zone Teams qui reçoit le
    // message ; puis envoyer() (outlook.js:723) retombe sur premier(SEL.envoyer) dans tout le document et clique
    // l'« Envoyer » de Teams ; outlook.js:727 (« plus aucune rédaction visible ») le compte comme envoyé (E6, E3, E1).
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.locator("#btn-teams").click(); await page.locator("#teams-saisie").click(); await page.keyboard.type("Je te rappelle demain");
    await page.evaluate(() => (window as any).ow.piege({ type: "input", zone: /^objet-/, action: () => (window as any).ow.afficherMessage(1) }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.teamsEnvoyes).toEqual([]);
    expect(e.teams).toBe("Je te rappelle demain");
    expect(await nbEnvoyeSelonSuivi(page)).toBe(e.envoyes.length); // le suivi ne compte que ce qui est vraiment parti
  });

  test("D04b … juste avant « Envoyer » (corps déjà collé), zone Teams sans aria-multiline : le bouton « Envoyer » de Teams n'est jamais cliqué", async ({ page, worker }) => {
    // BUG : publipostage/outlook.js:723 : rédaction masquée, zoneRedaction() = null, repli premier(SEL.envoyer) sur TOUT le
    // document → clic sur l'« Envoyer » de Teams (brouillon Teams de l'utilisateur envoyé) ; outlook.js:727 note le mail
    // comme « envoyé » alors qu'il reste dans un onglet masqué : la destinataire ne le recevra jamais (E6, E1/E2).
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { const ow = (window as any).ow; ow.afficherMessage(null); ow.teamsSansMultiligne(); });
    await page.locator("#btn-teams").click(); await page.locator("#teams-saisie").click(); await page.keyboard.type("Je te rappelle demain");
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^corps-/, action: () => (window as any).ow.afficherMessage(1) }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.teamsEnvoyes).toEqual([]);
    expect(await nbEnvoyeSelonSuivi(page)).toBe(e.envoyes.length);
  });

  test("D05 même chose sans Teams : arrêt clair, rien envoyé, rien compté à tort", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^corps-/, action: () => (window as any).ow.afficherMessage(1) }));
    await lancerEnvoi(page, { dest: DEST });
    const n = await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(await nbEnvoyeSelonSuivi(page)).toBe(0);
    expect(n).toMatch(/Arrêt avant l'envoi/);
  });

  test("D06 l'utilisateur abandonne la rédaction de l'envoi aussitôt ouverte, son brouillon redevient actif : rien n'y est tapé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.apresFermeture = "onglet-suivant"; });
    await brouillonUtilisateur(page);
    await page.evaluate(() => (window as any).ow.afficherMessage(2));
    const avant = contenuUtilisateur(await etat(page));
    await page.evaluate(() => (window as any).ow.piege({ type: "click", si: (e: MouseEvent) => (e.target as HTMLElement).id === "nouveau-courrier",
      action: () => { const ow = (window as any).ow; const e = ow.etat(); ow.fermer(e.redactions.find((r: any) => r.proprietaire === "extension").id); } }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(contenuUtilisateur(e)).toEqual(avant);
  });

  test("D07 abandon juste après le collage dans « À », brouillon de l'utilisateur (1 destinataire, objet) redevient actif : son objet n'est pas modifié", async ({ page, worker }) => {
    // BUG : même cause que D01/D02 : remplirA valide sur la pastille du brouillon de l'utilisateur, puis remplirObjet
    // (outlook.js:476-486) colle « Invitation » dans SON objet avant de constater l'écart (R5 : brouillon modifié).
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.apresFermeture = "onglet-suivant"; });
    await brouillonUtilisateur(page, { pastilles: ["celine.moreau@exemple.fr"], objet: "Budget 2027" });
    await page.evaluate(() => (window as any).ow.afficherMessage(2));
    const avant = contenuUtilisateur(await etat(page));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^a-/, action: () => { const ow = (window as any).ow; const e = ow.etat(); ow.fermer(e.redactions.find((r: any) => r.proprietaire === "extension").id); } }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(contenuUtilisateur(e)).toEqual(avant);
  });

  test("D08 même abandon, brouillon de l'utilisateur SANS objet (1 destinataire) : il n'est jamais rempli ni envoyé", async ({ page, worker }) => {
    // BUG : même cause : le brouillon de l'utilisateur reçoit l'objet et le message de la campagne et part à SON destinataire.
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.apresFermeture = "onglet-suivant"; });
    await brouillonUtilisateur(page, { pastilles: ["celine.moreau@exemple.fr"], objet: "", corps: "Céline, peux-tu relire ?" });
    await page.evaluate(() => (window as any).ow.afficherMessage(2));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^a-/, action: () => { const ow = (window as any).ow; const e = ow.etat(); ow.fermer(e.redactions.find((r: any) => r.proprietaire === "extension").id); } }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    const e = await etat(page);
    expect(e.envoyes.filter((m) => m.proprietaire === "utilisateur").map((m) => ({ a: m.a, objet: m.objet }))).toEqual([]);
  });

  test("D09 abandon juste après le collage du corps, brouillon de l'utilisateur redevient actif : jamais envoyé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.apresFermeture = "onglet-suivant"; });
    await brouillonUtilisateur(page, { pastilles: ["celine.moreau@exemple.fr"] });
    await page.evaluate(() => (window as any).ow.afficherMessage(2));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^corps-/, action: () => { const ow = (window as any).ow; const e = ow.etat(); ow.fermer(e.redactions.find((r: any) => r.proprietaire === "extension").id); } }));
    await lancerEnvoi(page, { dest: DEST });
    await attendreFinEnvoi(page);
    expect((await etat(page)).envoyes.filter((m) => m.proprietaire === "utilisateur")).toEqual([]);
  });

  test("D10 Outlook passe au Calendrier en cours de mail : arrêt clair, rien envoyé", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "input", zone: /^objet-/, action: () => (window as any).ow.vue("calendrier") }));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toMatch(/Arrêt avant l'envoi/);
    expect((await etat(page)).envoyes).toEqual([]);
  });

  test("D11 champ « À » re-rendu juste après le collage : un seul destinataire, le bon, envoi normal", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^a-/, action: () => { const ow = (window as any).ow; ow.rerendreA(ow.etat().actif); } }));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    verifierMailsExtension(await etat(page), ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]);
  });

  test("D12 une notification prend le focus quand l'extension clique l'objet : objet exact, rien tapé dans la notification", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "mousedown", zone: /^objet-/, action: () => (window as any).ow.notifier("saisie") }));
    await lancerEnvoi(page, { dest: DEST });
    expect(await attendreFinEnvoi(page)).toContain("Terminé : 2 mail(s) envoyé(s)");
    verifierMailsExtension(await etat(page), ["alice.martin@exemple.fr", "bruno.faure@exemple.fr"]);
    expect((await etat(page)).toast).toBe("");
  });

  test("D13 pendant la pause entre deux mails, l'utilisateur ouvre une réponse : le mail suivant est refusé, rien tapé dans la réponse", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(0));
    await lancerEnvoi(page, { dest: DEST, pmin: 300, pmax: 300 }); // ≈ 3 s de pause en mode accéléré
    await expect.poll(async () => (await etat(page)).envoyes.length, { timeout: 20000 }).toBe(1);
    await page.locator('#lecture [aria-label="Répondre"]').click();
    await page.keyboard.type("Merci Paul");
    const n = await attendreFinEnvoi(page);
    expect(n).toMatch(/déjà ouvert en rédaction/);
    const e = await etat(page);
    expect(e.envoyes.length).toBe(1);
    expect(e.redactions.find((r) => r.proprietaire === "utilisateur")!.corps).not.toContain("Voici l'invitation");
  });

  test("D14 (vitesse réelle) l'utilisateur tape dans la recherche pendant que l'extension colle le corps : ses frappes ne partent jamais dans le mail", async ({ page, worker }) => {
    // BUG : vol du focus : remplirCorps (outlook.js:548-553) donne le focus au corps et rien ne revérifie objet/corps
    // juste avant « Envoyer » (moteur.js:304-313) : les frappes de l'utilisateur, détournées dans le corps, partent
    // avec le mail (E3), et sa recherche est amputée.
    test.setTimeout(150000);
    await ouvrir(page, worker, "envoi", URL_OUTLOOK + "?vitesse=reelle");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await lancerEnvoi(page, { dest: "alice.martin@exemple.fr", pmin: 1, pmax: 1 });
    const ext = async () => (await etat(page)).redactions.find((r) => r.proprietaire === "extension");
    await expect.poll(async () => (await ext())?.objet ?? "", { timeout: 60000 }).toBe("Invitation");
    await page.locator("#recherche").click(); // l'utilisateur va chercher un mail…
    await page.keyboard.type("rapport", { delay: 60 }); // … et tape dans la recherche
    // Tant qu'il écrit ailleurs, l'extension ne lui prend pas le curseur ; elle reprend quand il clique ailleurs.
    await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur()); // il quitte la recherche
    await attendreFinEnvoi(page, 90000);
    const e = await etat(page);
    expect(e.envoyes.map((m) => m.corps.includes("rapport"))).toEqual(e.envoyes.map(() => false));
    expect(e.recherche).toBe("rapport");
  });
});

test.describe("Envoi — ⚙ « Lancer le diagnostic » avec un brouillon de l'utilisateur ouvert", () => {
  test("C14 diagnostic du Publipostage, brouillon de l'utilisateur ouvert : aucune commande cliquée dans ce brouillon", async ({ page, worker }) => {
    // BUG : publipostage/outlook.js:746-803 diagnostic() n'ouvre un nouveau message que si AUCUNE rédaction n'est
    // ouverte ; sinon il travaille dans celle de l'utilisateur : onglets du ruban, importanceHaute() (clic « Importance
    // élevée » sur SON brouillon), puis « Ignorer » s'il existe.
    await ouvrir(page, worker, "envoi");
    await brouillonUtilisateur(page);
    await pp(page).locator("#pp-gear").click();
    const dl = page.waitForEvent("download", { timeout: 30000 }).catch(() => null);
    await pp(page).locator("#pp-diag").click();
    await dl;
    await page.waitForTimeout(300);
    const clics = (await clicsExtension(page)).map((c) => c.libelle);
    expect(clics).toEqual([]);
  });

  test("C15 même diagnostic, variante « Ignorer » (ancien OWA FR) : le brouillon de l'utilisateur n'est jamais ignoré", async ({ page, worker }) => {
    // BUG : outlook.js:795-800 clique premier(SEL.ignorer) (« Ignorer »/« Discard », outlook.js:46) puis confirme :
    // c'est le brouillon de l'UTILISATEUR qui est jeté.
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.libelleAbandon = "Ignorer"; });
    await brouillonUtilisateur(page);
    await pp(page).locator("#pp-gear").click();
    const dl = page.waitForEvent("download", { timeout: 30000 }).catch(() => null);
    await pp(page).locator("#pp-diag").click();
    await dl;
    await page.waitForTimeout(300);
    expect((await etat(page)).redactions.filter((r) => r.proprietaire === "utilisateur").map((r) => r.objet)).toEqual(["Budget 2027"]);
  });

  test("C16 diagnostic du Publipostage depuis la boîte de réception : ouvre SA rédaction, ne touche à rien d'autre, la referme", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => { (window as any).ow.reglages.libelleAbandon = "Ignorer"; (window as any).ow.afficherMessage(null); });
    await pp(page).locator("#pp-gear").click();
    const dl = page.waitForEvent("download", { timeout: 30000 }).catch(() => null);
    await pp(page).locator("#pp-diag").click();
    await dl;
    await page.waitForTimeout(300);
    const e = await etat(page);
    expect(e.envoyes).toEqual([]);
    expect(e.redactions.filter((r) => r.proprietaire === "utilisateur")).toEqual([]);
  });

  test("D06b la rédaction de l'envoi est abandonnée juste après le collage du corps : arrêt clair, rien envoyé ni compté", async ({ page, worker }) => {
    await ouvrir(page, worker, "envoi");
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "paste", zone: /^corps-/, action: () => { const ow = (window as any).ow; ow.fermer(ow.etat().actif); } }));
    await lancerEnvoi(page, { dest: DEST });
    const n = await attendreFinEnvoi(page);
    expect((await etat(page)).envoyes).toEqual([]);
    expect(await nbEnvoyeSelonSuivi(page)).toBe(0);
    expect(n).toMatch(/Arrêt avant l'envoi/);
  });
});
