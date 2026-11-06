// Audit Smart Mail — « Champ visé, brouillons de l'utilisateur et vol du focus » (2/3) :
// PENDANT un relevé (lancé depuis la boîte de réception, dans la rédaction ouverte par l'extension) :
// l'utilisateur clique ailleurs, une notification prend le focus, la rédaction est fermée, une autre s'ouvre,
// Outlook change de vue, le champ « À » est re-rendu, le focus bouge pendant un effacement.
// Règles : R4 (jamais de frappe hors du champ visé ; s'il perd le focus : arrêt propre), R5, R1, R2, R6.
import { test, expect, ao, ouvrir, etat, journal, saisiesExtension, saisiesUtilisateur, demarrerReleve, arreterReleve,
  attendreRecherches, noticeReleve, releveEnCours, releveLance, reglagesReleve, attendreFocus, validations, viderJournal } from "./audit/champ-cible-brouillons/banc.ts";
import type { Page } from "@playwright/test";

test.describe.configure({ mode: "parallel" });

// Zones où l'extension a tapé/effacé, hors du « À » de SA rédaction.
async function horsCible(page: Page, id: string) {
  return (await saisiesExtension(page)).filter((s) => s.zone !== "a-" + id).map((s) => `${s.zone}:${s.inputType}:${s.data ?? ""}`);
}
// Le relevé a-t-il rendu la main (vue compacte fermée) dans le délai ?
async function arreteDans(page: Page, ms: number) {
  const fin = Date.now() + ms;
  while (Date.now() < fin) { if (!(await releveEnCours(page))) return true; await page.waitForTimeout(100); }
  return false;
}
async function pauseSiBesoin(page: Page) { if (await releveEnCours(page)) await arreterReleve(page); }
async function attendreDemarrage(page: Page) { await expect(ao(page).locator("#ao-compact-view")).toBeVisible({ timeout: 20000 }); }

test.describe("Relevé en cours — l'utilisateur ou Outlook prend le focus", () => {
  test("B01 l'utilisateur clique dans la recherche et tape « paul » : le relevé s'arrête et ne reprend pas le focus", async ({ page, worker }) => {
    // BUG : adapter-dom.js:130-136 focus() appelle el.focus() AVANT de vérifier : le moteur (core/moteur.js:504)
    // ne voit jamais que le curseur a quitté « À » ; il reprend le focus à la recherche suivante, sans s'arrêter.
    await releveLance(page, worker);
    await page.locator("#recherche").click();
    await page.keyboard.type("paul", { delay: 120 });
    const vole = await attendreFocus(page, /^a-/, 6000);
    const arrete = !(await releveEnCours(page));
    await pauseSiBesoin(page);
    expect({ focusRepris: vole, releveArrete: arrete }).toEqual({ focusRepris: false, releveArrete: true });
    expect((await etat(page)).recherche).toBe("paul");
  });

  test("B02 … puis Entrée pour lancer sa recherche : aucun destinataire n'est jamais validé (R1)", async ({ page, worker }) => {
    // BUG : conséquence de B01 : le focus repris, l'Entrée de l'utilisateur part dans « À » et valide une suggestion.
    const id = await releveLance(page, worker);
    await page.locator("#recherche").click();
    await page.keyboard.type("paul", { delay: 120 });
    if (await attendreFocus(page, /^a-/, 6000)) {
      // l'utilisateur appuie sur Entrée quand la liste de suggestions du relevé est affichée
      for (let i = 0; i < 80 && !(await page.locator("#suggestions [role=option]").count()); i++) await page.waitForTimeout(100);
    }
    await page.keyboard.press("Enter");
    await page.waitForTimeout(500);
    await pauseSiBesoin(page);
    // Une pastille validée puis effacée par le relevé lui-même (il ne la voit pas) reste une validation (R1).
    expect((await validations(page)).filter((v) => v.zone === "a-" + id)).toEqual([]);
  });

  test("B03 une notification prend le focus (bouton « Afficher ») : arrêt explicite, focus laissé à la notification", async ({ page, worker }) => {
    // BUG : même cause que B01 (adapter-dom.js:130-136) : le focus est repris en silence, le relevé continue.
    await releveLance(page, worker);
    await page.evaluate(() => (window as any).ow.notifier("bouton"));
    const vole = await attendreFocus(page, /^a-/, 6000);
    const arrete = !(await releveEnCours(page));
    const notice = await noticeReleve(page);
    await pauseSiBesoin(page);
    expect({ focusRepris: vole, releveArrete: arrete }).toEqual({ focusRepris: false, releveArrete: true });
    expect(notice).toMatch(/curseur|focus|champ/i);
  });

  test("B04 l'utilisateur clique dans le corps du message du relevé et écrit « Bonjour » : ses lettres ne partent jamais dans « À »", async ({ page, worker }) => {
    // BUG : même cause que B01 : le moteur reprend le focus en cours de frappe, les lettres suivantes vont dans « À ».
    const id = await releveLance(page, worker);
    await page.locator(`[data-zone="corps-${id}"] .elementToProof`).first().click();
    await page.keyboard.type("Bonjour Paul, merci", { delay: 250 }); // ~5 s : au moins une recherche passe pendant la frappe
    await pauseSiBesoin(page);
    const u = await saisiesUtilisateur(page);
    expect(u.filter((x) => x.zone !== "corps-" + id).map((x) => `${x.zone}:${x.data}`)).toEqual([]);
    expect((await validations(page)).filter((v) => v.zone === "a-" + id)).toEqual([]); // la virgule de l'utilisateur ne valide rien
  });

  test("B05 modale « session expirée » (champ mot de passe focalisé) pendant le relevé : arrêt, focus non repris", async ({ page, worker }) => {
    // BUG : même cause que B01 ; variante plausible d'une fenêtre de reconnexion (E5/R4) : le relevé continue dessous.
    await releveLance(page, worker);
    await page.evaluate(() => {
      const d = document.createElement("div"); d.setAttribute("role", "dialog"); d.setAttribute("aria-modal", "true"); d.className = "ms-Modal is-open"; d.id = "session";
      d.innerHTML = '<p>Votre session a expiré. Entrez votre mot de passe.</p><input type="password" aria-label="Mot de passe" data-zone="mdp" id="mdp"><button type="button">Se connecter</button>';
      document.getElementById("portail")!.appendChild(d); (document.getElementById("mdp") as HTMLInputElement).focus();
    });
    const vole = await attendreFocus(page, /^a-/, 6000);
    const arrete = !(await releveEnCours(page));
    await pauseSiBesoin(page);
    expect({ focusRepris: vole, releveArrete: arrete }).toEqual({ focusRepris: false, releveArrete: true });
  });

  test("B06 le focus passe dans le brouillon Teams pendant un effacement : aucun caractère de Teams effacé", async ({ page, worker }) => {
    // BUG : adapter-dom.js:104-111 la boucle d'effacement appelle execCommand("delete") sur l'élément FOCALISÉ sans
    // re-vérifier le focus entre deux touches (pause de 25 ms) : si le focus bouge (clic, notification), elle efface ailleurs.
    // Recherches de 3 lettres dès le départ (préfixe « pl », réglage enregistré) pour ne pas attendre la phase d'approfondissement.
    await reglagesReleve(worker, { Prefix: "pl" });
    await ouvrir(page, worker);
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.locator("#btn-teams").click();
    await page.locator("#teams-saisie").click(); await page.keyboard.type("Je te rappelle demain matin", { delay: 5 });
    const avant = (await etat(page)).teams;
    await page.locator("#lecture .vide").click(); // l'utilisateur a laissé Teams : « Démarrer » ouvre la rédaction de l'extension
    // L'utilisateur clique à la fin de son message Teams pendant le 1er effacement d'une recherche de 3 lettres.
    await page.evaluate(() => (window as any).ow.piege({ type: "keydown", zone: /^a-/, cle: "Backspace", action: () => (window as any).ow.curseurFin("#teams-saisie") }));
    await demarrerReleve(page);
    await attendreRecherches(page, 3);
    await pauseSiBesoin(page);
    expect((await etat(page)).teams).toBe(avant);
  });

  test("B07 même course, focus dans la barre de recherche (« facture mars ») : la recherche n'est pas amputée", async ({ page, worker }) => {
    // BUG : même cause que B06 (adapter-dom.js:104-111).
    await reglagesReleve(worker, { Prefix: "pl" });
    await ouvrir(page, worker);
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.locator("#recherche").click(); await page.keyboard.type("facture mars");
    await page.locator("#lecture .vide").click();
    await page.evaluate(() => (window as any).ow.piege({ type: "keydown", zone: /^a-/, cle: "Backspace", action: () => (window as any).ow.curseurFin("#recherche") }));
    await demarrerReleve(page);
    await attendreRecherches(page, 3);
    await pauseSiBesoin(page);
    expect((await etat(page)).recherche).toBe("facture mars");
  });

  test("B08 même course SANS préfixe (réglages d'origine), pendant une recherche de 2 lettres : rien d'effacé dans Teams", async ({ page, worker }) => {
    // BUG : même cause que B06, avec les réglages par défaut : la phase d'approfondissement tape des recherches de 2 lettres.
    test.setTimeout(240000);
    await ouvrir(page, worker);
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.locator("#btn-teams").click();
    await page.locator("#teams-saisie").click(); await page.keyboard.type("Je te rappelle demain matin", { delay: 5 });
    const avant = (await etat(page)).teams;
    await page.locator("#lecture .vide").click();
    await page.evaluate(() => (window as any).ow.piege({ type: "keydown", zone: /^a-/, cle: "Backspace",
      si: (e: KeyboardEvent) => ((e.target as HTMLElement).textContent || "").replace(/​/g, "").length >= 2,
      action: () => (window as any).ow.curseurFin("#teams-saisie") }));
    await demarrerReleve(page);
    // Soit le relevé a tapé assez longtemps pour croiser le piège, soit il s'est arrêté proprement (curseur volé).
    await expect.poll(async () => (await saisiesExtension(page)).filter((s) => s.inputType === "insertText" && (s.data || "").length > 0).length > 29 || !(await releveEnCours(page)), { timeout: 200000 }).toBe(true);
    await pauseSiBesoin(page);
    expect((await etat(page)).teams).toBe(avant);
  });
});

test.describe("Relevé en cours — la rédaction disparaît ou change", () => {
  test("B09 l'utilisateur abandonne la rédaction du relevé : arrêt explicite, rien tapé ailleurs", async ({ page, worker }) => {
    const id = await releveLance(page, worker);
    await page.locator(`.redaction[data-redaction="${id}"] [aria-label="Abandonner"]`).click();
    expect(await arreteDans(page, 10000)).toBe(true);
    expect(await noticeReleve(page)).toMatch(/Arrêté|pause|curseur|champ/i);
    expect(await horsCible(page, id)).toEqual([]);
  });

  test("B10 … et le brouillon de l'utilisateur (destinataires, texte) redevient l'onglet actif : rien n'y est tapé", async ({ page, worker }) => {
    await ouvrir(page, worker);
    const u = await page.evaluate(() => { const ow = (window as any).ow; ow.reglages.apresFermeture = "onglet-suivant"; const id = ow.nouvelleRedaction("utilisateur"); ow.ajouterPastille(id, "À", "alice.martin@exemple.fr"); ow.ecrireObjet(id, "Budget"); ow.ecrireCorps(id, "Texte du budget."); ow.afficherMessage(1); return id; });
    const avant = (await etat(page)).redactions.find((x) => x.id === u)!;
    await demarrerReleve(page);
    await attendreRecherches(page, 2);
    const id = (await etat(page)).redactions.find((x) => x.proprietaire === "extension")!.id;
    await page.locator(`.redaction[data-redaction="${id}"] [aria-label="Abandonner"]`).click();
    expect(await arreteDans(page, 10000)).toBe(true);
    const apres = (await etat(page)).redactions.find((x) => x.id === u)!;
    expect(apres).toEqual({ ...avant, visible: true });
    expect(await horsCible(page, id)).toEqual([]);
  });

  test("B11 l'utilisateur ouvre un nouveau message pendant le relevé : rien n'y est tapé, arrêt explicite", async ({ page, worker }) => {
    const id = await releveLance(page, worker);
    await page.locator("#nouveau-courrier").click();
    expect(await arreteDans(page, 10000)).toBe(true);
    await page.waitForTimeout(1500);
    expect(await horsCible(page, id)).toEqual([]);
    expect(await noticeReleve(page)).not.toBe("");
  });

  test("B12 Outlook passe au Calendrier pendant le relevé : arrêt explicite, rien tapé dans le calendrier", async ({ page, worker }) => {
    const id = await releveLance(page, worker);
    await page.locator('nav#rail [data-vue="calendrier"]').click();
    await page.locator("#nouvel-evenement").click();
    await page.locator("#ev-titre").click();
    expect(await arreteDans(page, 10000)).toBe(true);
    await page.waitForTimeout(1500);
    expect(await horsCible(page, id)).toEqual([]);
  });

  test("B13 le champ « À » est re-rendu (nouvel élément) : aucune frappe hors de ce champ, aucune boucle", async ({ page, worker }) => {
    const id = await releveLance(page, worker);
    await page.evaluate((i) => (window as any).ow.rerendreA(i), id);
    await page.waitForTimeout(4000);
    await pauseSiBesoin(page);
    expect(await horsCible(page, id)).toEqual([]);
  });

  test("B14 R2 : interrompu pendant une recherche (onglet de l'utilisateur ouvert par-dessus), aucun texte du relevé ne reste dans « À »", async ({ page, worker }) => {
    // BUG : core/moteur.js:542 n'efface qu'« if (focus().ok) » ; le champ masqué n'est plus jamais nettoyé : quand
    // l'utilisateur revient sur l'onglet du relevé, la recherche tapée est toujours dans « À ».
    await ouvrir(page, worker);
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "input", zone: /^a-/, si: (e: InputEvent) => e.inputType === "insertText", apres: 3, action: () => (window as any).ow.nouvelleRedaction("utilisateur") }));
    await demarrerReleve(page);
    await attendreDemarrage(page);
    expect(await arreteDans(page, 30000)).toBe(true);
    const r = (await etat(page)).redactions.find((x) => x.proprietaire === "extension")!;
    expect(r.aTexte.replace(/\s/g, "")).toBe("");
  });

  test("B15 R2 : champ « À » re-rendu juste après une frappe : aucun texte du relevé ne reste dans le nouveau champ", async ({ page, worker }) => {
    // BUG : même cause que B14 : arrêt « le curseur a quitté le champ » sans nettoyer le champ qui le remplace.
    await ouvrir(page, worker);
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await page.evaluate(() => (window as any).ow.piege({ type: "input", zone: /^a-/, si: (e: InputEvent) => e.inputType === "insertText", apres: 3,
      action: () => { const ow = (window as any).ow; const e = ow.etat(); ow.rerendreA(e.redactions.find((x: any) => x.proprietaire === "extension").id); } }));
    await demarrerReleve(page);
    await attendreDemarrage(page);
    await page.waitForTimeout(9000);
    await pauseSiBesoin(page);
    const r = (await etat(page)).redactions.find((x) => x.proprietaire === "extension")!;
    expect(r.aTexte.replace(/\s/g, "")).toBe("");
  });
});
