// Audit Smart Mail — « Champ visé, brouillons de l'utilisateur et vol du focus » (1/3) :
// le bouton « Démarrer » du RELEVÉ depuis tous les états d'Outlook.
// Fausse page : tests/audit/champ-cible-brouillons/owa-monarch.html (nouvel Outlook FR fidèle à la fiche de
// référence ; ce qui n'est pas connu du vrai Outlook y est signalé comme variante plausible).
// Règles : R4 (jamais de frappe hors du champ destinataire visé), R5 (jamais d'effacement de ce que
// l'utilisateur a mis), R2, R1, R3. Un test qui échoue porte « // BUG : … ».
// Lancer : node node_modules/@playwright/test/cli.js test tests/audit-champ-cible-brouillons.spec.ts --workers=2 --reporter=line
import { test, expect, ao, ouvrir, etat, saisiesExtension, demarrerReleve, attendreDebutOuRefus, arreterReleve,
  attendreRecherches, noticeReleve, viderJournal, releveEnCours, contenuUtilisateur, type Etat } from "./audit/champ-cible-brouillons/banc.ts";
import type { Page, Worker } from "@playwright/test";

test.describe.configure({ mode: "parallel" }); // chaque test a son propre profil Chromium : indépendants

type Resultat = { issue: string; avant: Etat; apres: Etat; horsCible: string[]; notice: string; lus: string[] };

// Lance « Démarrer » après une préparation, laisse le relevé travailler `apres` ms, le met en pause, et rend :
// les saisies de l'extension HORS du « À » d'une rédaction qu'elle a ouverte elle-même (ou d'une zone permise).
async function releveDepuis(page: Page, worker: Worker, preparer: (p: Page) => Promise<void>, o: { apres?: number; permises?: string[] } = {}): Promise<Resultat> {
  await ouvrir(page, worker);
  await preparer(page);
  const avant = await etat(page);
  await viderJournal(page);
  await demarrerReleve(page);
  const issue = await attendreDebutOuRefus(page, 25000);
  if (issue === "saisie") await page.waitForTimeout(o.apres ?? 2500);
  if (await releveEnCours(page)) await arreterReleve(page);
  await page.waitForTimeout(300);
  const apres = await etat(page);
  const permises = new Set([...(o.permises ?? []), ...apres.redactions.filter((r) => r.proprietaire === "extension").map((r) => "a-" + r.id)]);
  const horsCible = (await saisiesExtension(page)).filter((s) => !permises.has(s.zone)).map((s) => `${s.zone}:${s.inputType}:${s.data ?? ""}`);
  const lus = await ao(page).locator("#ao-rows .m").allTextContents();
  return { issue, avant, apres, horsCible, notice: await noticeReleve(page), lus };
}
// Résumé lisible : zones touchées et nombre de saisies par zone (pour le message d'échec).
const resume = (l: string[]) => Object.entries(l.reduce((m: Record<string, number>, s) => { const k = s.split(":").slice(0, 2).join(":"); m[k] = (m[k] || 0) + 1; return m; }, {})).map(([k, n]) => `${k}×${n}`);

// Un brouillon de l'utilisateur : réponse en ligne au message 0 (pastille de l'expéditeur + une ajoutée, texte écrit).
const TEXTE_REPONSE = "Bonjour Paul, voici le planning de la semaine prochaine, je te confirme jeudi.";
async function reponseUtilisateur(page: Page) {
  return page.evaluate((t) => { const ow = (window as any).ow; const id = ow.repondre(0, "utilisateur"); ow.ajouterPastille(id, "À", "claire.petit@exemple.fr"); ow.ecrireCorps(id, t); return id as string; }, TEXTE_REPONSE);
}
// Un nouveau message de l'utilisateur, en cours d'écriture : deux destinataires, un objet, un texte.
async function brouillonUtilisateur(page: Page) {
  return page.evaluate(() => {
    const ow = (window as any).ow; const id = ow.nouvelleRedaction("utilisateur");
    ow.ajouterPastille(id, "À", "alice.martin@exemple.fr"); ow.ajouterPastille(id, "À", "bruno.faure@exemple.fr");
    ow.ecrireObjet(id, "Budget 2027"); ow.ecrireCorps(id, "Bonjour à tous, vous trouverez ci-joint le budget prévisionnel."); return id as string;
  });
}
// Vrai clic dans le champ puis Fin : le curseur est à la fin du texte, comme quand on écrit. On vérifie que le
// focus est bien dans la zone voulue (sinon le scénario ne serait pas celui annoncé).
async function cliquerFinDe(page: Page, sel: string) {
  const el = page.locator(sel).first();
  await el.click();
  await page.keyboard.press("End");
  const zone = await el.evaluate((x) => (x.closest("[data-zone]") as HTMLElement | null)?.dataset.zone ?? "");
  expect((await etat(page)).actifZone).toBe(zone);
}

test.describe("Relevé — « Démarrer » depuis les états d'Outlook", () => {
  test("A01 boîte de réception, volet vide : nouveau message de l'extension, frappe uniquement dans son « À »", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.evaluate(() => (window as any).ow.afficherMessage(null)); });
    expect(r.issue).toBe("saisie");
    expect(r.apres.redactions.map((x) => x.proprietaire)).toEqual(["extension"]);
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.redactions[0].aTexte.replace(/\s/g, "")).toBe(""); // R2
    expect(r.apres.redactions[0].a).toEqual([]);                      // R1
  });

  test("A02 message lu dans le volet de lecture : même chose, le message lu n'est pas touché", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.evaluate(() => (window as any).ow.afficherMessage(1)); await p.locator("#lecture .texte-lu").click(); });
    expect(r.issue).toBe("saisie");
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.redactions.map((x) => x.proprietaire)).toEqual(["extension"]);
  });

  test("A03 réponse en ligne de l'utilisateur, curseur dans son corps : ni frappe ni effacement dans ce corps", async ({ page, worker }) => {
    // BUG : adresses/panel.js:465-471 champActif() accepte le « Corps du message » (seuls search/objet/subject sont
    // écartés) ; armement.js:138-145 armHere l'arme, et le 1er effacement du passage (adapter-dom.js:103, n ≥ 200)
    // efface le texte de l'utilisateur avant le curseur, puis le moteur tape ses lettres dans le corps.
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => { id = await reponseUtilisateur(p); await cliquerFinDe(p, `[data-zone="corps-${id}"] .elementToProof`); });
    expect(resume(r.horsCible)).toEqual([]);
    expect(contenuUtilisateur(r.apres)).toEqual(contenuUtilisateur(r.avant));
  });

  test("A04 réponse en ligne, curseur à la fin de son « À » (2 pastilles) : aucune pastille effacée", async ({ page, worker }) => {
    // BUG : 1er effacement du passage n ≥ 200 (adapter-dom.js:103) et fieldEmpty (adapter-dom.js:98) compte le texte
    // des pastilles : la boucle d'effacement (adapter-dom.js:104-111) mange les destinataires de l'utilisateur.
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => { id = await reponseUtilisateur(p); await cliquerFinDe(p, `[data-zone="a-${id}"]`); });
    const avant = r.avant.redactions.find((x) => x.id === id)!, apres = r.apres.redactions.find((x) => x.id === id)!;
    expect(apres.a).toEqual(avant.a);
    expect(resume(r.horsCible)).toEqual([]);
  });

  test("A05 réponse en ligne ouverte, curseur sur le message lu (hors champ) : rien n'est tapé dans la réponse de l'utilisateur", async ({ page, worker }) => {
    // BUG : adresses/panel.js:479 prend O.champA() = le premier « À » visible (publipostage/outlook.js:146), donc celui
    // du brouillon de l'utilisateur, sans vérifier qu'il est vide ni qu'il appartient à l'extension (R4).
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => { id = await reponseUtilisateur(p); await p.locator("#lecture .texte-lu").click(); });
    expect(resume(r.horsCible)).toEqual([]);
    expect(contenuUtilisateur(r.apres)).toEqual(contenuUtilisateur(r.avant));
  }, );

  test("A05b même réponse ouverte, variante « À » qui met le curseur en fin de champ au focus : aucune pastille effacée", async ({ page, worker }) => {
    // BUG : panel.js:479 (champA du brouillon de l'utilisateur) + 1er effacement n ≥ 200 (adapter-dom.js:103-111) :
    // si Outlook place le curseur après les pastilles, les destinataires de l'utilisateur sont effacés sans qu'il ait cliqué.
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => {
      await p.evaluate(() => { (window as any).ow.reglages.curseurFinAuFocus = true; });
      id = await reponseUtilisateur(p); await p.locator("#lecture .texte-lu").click();
    });
    const avant = r.avant.redactions.find((x) => x.id === id)!, apres = r.apres.redactions.find((x) => x.id === id)!;
    expect(apres.a).toEqual(avant.a);
  });

  test("A06 brouillon de l'utilisateur (À, objet, corps), curseur dans son objet : rien n'est tapé dans ce brouillon", async ({ page, worker }) => {
    // BUG : champActif() écarte l'objet, puis panel.js:479 retombe sur O.champA() = « À » du brouillon de l'utilisateur.
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => { id = await brouillonUtilisateur(p); await cliquerFinDe(p, `[data-zone="objet-${id}"]`); });
    expect(resume(r.horsCible)).toEqual([]);
    expect(contenuUtilisateur(r.apres)).toEqual(contenuUtilisateur(r.avant));
  });

  test("A07 brouillon de l'utilisateur, curseur dans son corps : ni frappe ni effacement", async ({ page, worker }) => {
    // BUG : même cause que A03 (panel.js:465-471 + adapter-dom.js:103).
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => { id = await brouillonUtilisateur(p); await cliquerFinDe(p, `[data-zone="corps-${id}"] .elementToProof`); });
    expect(resume(r.horsCible)).toEqual([]);
    expect(contenuUtilisateur(r.apres)).toEqual(contenuUtilisateur(r.avant));
  });

  test("A08 brouillon : destinataire en cours de saisie « jean.d » après 2 pastilles, curseur dans « À » : rien d'effacé", async ({ page, worker }) => {
    // BUG : le « 1er effacement qui vide ENTIÈREMENT » (adapter-dom.js:101-103) date de l'armement par frappe ; avec
    // « Démarrer » il efface la saisie de l'utilisateur ET ses pastilles (R5).
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => {
      id = await brouillonUtilisateur(p);
      await cliquerFinDe(p, `[data-zone="a-${id}"]`);
      await p.keyboard.type("jean.d", { delay: 30 }); // l'utilisateur commençait à taper un destinataire
      await p.keyboard.press("Escape");
    });
    const avant = r.avant.redactions.find((x) => x.id === id)!, apres = r.apres.redactions.find((x) => x.id === id)!;
    expect({ a: apres.a, texte: apres.aTexte }).toEqual({ a: avant.a, texte: avant.aTexte });
  });

  test("A09 deux rédactions en onglets, l'active est un brouillon de l'utilisateur : rien n'est tapé dedans", async ({ page, worker }) => {
    // BUG : panel.js:479 O.champA() = « À » de l'onglet actif, donc du brouillon de l'utilisateur.
    const r = await releveDepuis(page, worker, async (p) => {
      await brouillonUtilisateur(p);
      await p.evaluate(() => { const ow = (window as any).ow; const id = ow.nouvelleRedaction("utilisateur"); ow.ajouterPastille(id, "À", "celine.moreau@exemple.fr"); ow.ecrireObjet(id, "Réunion"); });
      await p.locator('[role="tablist"][aria-label="Brouillons ouverts"] [role="tab"]').last().click(); // l'utilisateur est sur son onglet
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(contenuUtilisateur(r.apres)).toEqual(contenuUtilisateur(r.avant));
  });

  test("A10 brouillon de l'utilisateur dans un onglet masqué, message lu affiché : l'extension ouvre sa rédaction, brouillon intact", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await brouillonUtilisateur(p); await p.evaluate(() => (window as any).ow.afficherMessage(2)); });
    expect(r.issue).toBe("saisie");
    expect(resume(r.horsCible)).toEqual([]);
    expect(contenuUtilisateur(r.apres)).toEqual(contenuUtilisateur(r.avant));
  });

  test("A11 Calendrier (sans formulaire) : refus explicite, rien tapé nulle part", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.locator('nav#rail [data-vue="calendrier"]').click(); });
    expect(r.issue).toBe("refus");
    expect(r.notice).toMatch(/Nouveau message|introuvable|Ouvrez/i);
    expect(resume(r.horsCible)).toEqual([]);
  });

  test("A12 Calendrier, curseur dans le titre d'un événement : ni frappe ni effacement du titre", async ({ page, worker }) => {
    // BUG : champActif() (panel.js:465-471) accepte n'importe quel INPUT éditable (« Ajouter un titre »).
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator('nav#rail [data-vue="calendrier"]').click(); await p.locator("#nouvel-evenement").click();
      await p.locator("#ev-titre").click(); await p.keyboard.type("Point budget", { delay: 20 });
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.cal).toEqual(r.avant.cal);
  });

  test("A13 Calendrier, curseur dans « Inviter des participants » (1 participant) : rien d'effacé, rien tapé", async ({ page, worker }) => {
    // BUG : même cause (champActif accepte tout champ éditable) ; le participant de l'utilisateur est effacé.
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator('nav#rail [data-vue="calendrier"]').click(); await p.locator("#nouvel-evenement").click();
      await p.locator("#ev-participants").click(); await p.keyboard.type("paul.durand@exemple.fr;", { delay: 10 });
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.cal).toEqual(r.avant.cal);
  });

  test("A14 Calendrier, curseur dans la description de l'événement : ni frappe ni effacement", async ({ page, worker }) => {
    // BUG : même cause (panel.js:465-471, adapter-dom.js:103).
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator('nav#rail [data-vue="calendrier"]').click(); await p.locator("#nouvel-evenement").click();
      await p.locator("#ev-description").click(); await p.keyboard.type("Ordre du jour : budget et planning", { delay: 10 });
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.cal).toEqual(r.avant.cal);
  });

  test("A15 Contacts, curseur dans « Prénom » d'un nouveau contact : ni frappe ni effacement", async ({ page, worker }) => {
    // BUG : même cause (champActif accepte tout INPUT éditable).
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator('nav#rail [data-vue="contacts"]').click(); await p.locator("#nouveau-contact").click();
      await p.locator("#ct-prenom").click(); await p.keyboard.type("Jean-Baptiste", { delay: 10 });
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.contact).toEqual(r.avant.contact);
  });

  test("A16 Paramètres (fenêtre modale), curseur dans l'éditeur de signature : ni frappe ni effacement", async ({ page, worker }) => {
    // BUG : même cause ; la signature de l'utilisateur est effacée/modifiée.
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator("#btn-parametres").click();
      await cliquerFinDe(p, "#param-signature");
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.signature).toBe(r.avant.signature);
  });

  test("A17 Paramètres, curseur dans « Rechercher dans les paramètres » : refus explicite, rien tapé", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.locator("#btn-parametres").click(); await p.locator("#param-recherche").click(); await p.keyboard.type("signature"); });
    expect(r.issue).toBe("refus");
    expect(r.notice).toMatch(/fenêtre|Fermez/i);
    expect(resume(r.horsCible)).toEqual([]);
  });

  test("A18 résultats de recherche affichés, curseur dans la recherche : recherche intacte, frappe seulement dans le « À » de l'extension", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.locator("#recherche").click(); await p.keyboard.type("planning"); await p.keyboard.press("Enter"); });
    expect(r.issue).toBe("saisie");
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.recherche).toBe("planning");
  });

  test("A19 résultats de recherche affichés : aucune adresse lue dans les résultats de recherche (R3)", async ({ page, worker }) => {
    // BUG : adapter-dom.js:160-161 resolveListbox prend la PREMIÈRE [role=listbox] visible contenant une adresse ;
    // isMsgList (adapter-dom.js:145-148) n'écarte que « liste des messages / courrier / boîte de réception / inbox » :
    // la liste « Résultats de la recherche » (variante plausible) est lue comme des suggestions.
    const r = await releveDepuis(page, worker, async (p) => { await p.locator("#recherche").click(); await p.keyboard.type("planning"); await p.keyboard.press("Enter"); }, { apres: 4000 });
    expect(r.lus).not.toContain("infolettre@diffusion-exemple.com"); // adresse présente SEULEMENT dans un aperçu de résultat
  });

  test("A20 curseur dans la barre de recherche (texte tapé, pas lancé) : texte intact, rien tapé dedans", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.locator("#recherche").click(); await p.keyboard.type("facture mars"); });
    expect(r.issue).toBe("saisie");
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.recherche).toBe("facture mars");
  });

  test("A21 nouveau message vide de l'utilisateur, curseur dans l'objet déjà tapé : objet intact, frappe dans son « À » vide seulement", async ({ page, worker }) => {
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator("#nouveau-courrier").click(); // vrai clic : rédaction de l'utilisateur
      id = (await etat(p)).actif!;
      await p.locator(`[data-zone="objet-${id}"]`).click(); await p.keyboard.type("Réunion jeudi");
    }, { permises: ["a-c1"] });
    // Choix de prudence : un message où l'utilisateur a déjà tapé un objet est le sien,
    // le relevé refuse d'y démarrer (refus) — dans tous les cas rien n'est touché.
    expect(["saisie", "refus"]).toContain(r.issue);
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.redactions.find((x) => x.id === id)!.objet).toBe("Réunion jeudi");
  });

  test("A22 nouveau message vide de l'utilisateur, curseur dans le corps : jamais de lettres tapées dans le corps", async ({ page, worker }) => {
    // BUG : même cause que A03 (champActif accepte « Corps du message ») : le relevé tape dans le corps.
    let id = "";
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator("#nouveau-courrier").click();
      id = (await etat(p)).actif!;
      await p.locator(`[data-zone="corps-${id}"] .elementToProof`).first().click();
    }, { permises: ["a-c1"] });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.redactions.find((x) => x.id === id)!.corps).toBe(r.avant.redactions.find((x) => x.id === id)!.corps);
  });

  test("A23 discussion Teams intégrée, curseur dans « Tapez un message » (brouillon) : ni frappe ni effacement", async ({ page, worker }) => {
    // BUG : même cause (champActif accepte la zone de saisie Teams, variante plausible dans le DOM principal).
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator("#btn-teams").click();
      await p.locator("#teams-saisie").click(); await p.keyboard.type("Oui je regarde le planning", { delay: 10 });
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.teams).toBe(r.avant.teams);
    expect(r.apres.teamsEnvoyes).toEqual([]);
  });

  test("A24 notification : curseur dans sa réponse rapide (texte tapé) : ni frappe ni effacement", async ({ page, worker }) => {
    // BUG : même cause (champActif accepte tout INPUT hors recherche/objet).
    const r = await releveDepuis(page, worker, async (p) => {
      await p.evaluate(() => (window as any).ow.notifier("saisie"));
      await p.keyboard.type("oui dans 5 min", { delay: 10 });
    });
    expect(resume(r.horsCible)).toEqual([]);
    expect(r.apres.toast).toBe(r.avant.toast);
  });

  test("A25 nouveau message vide de l'utilisateur, curseur dans « Cc » : le relevé tape dans ce Cc vide, rien d'autre", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => {
      await p.locator("#nouveau-courrier").click();
      const id = (await etat(p)).actif!;
      await p.locator(`[data-zone="cc-${id}"]`).click();
    }, { permises: ["cc-c1"] });
    expect(r.issue).toBe("saisie");
    expect(resume(r.horsCible)).toEqual([]);
  });

  test("A26 To Do (vue sans courrier) : refus explicite, rien tapé", async ({ page, worker }) => {
    const r = await releveDepuis(page, worker, async (p) => { await p.locator('nav#rail [data-vue="todo"]').click(); });
    expect(r.issue).toBe("refus");
    expect(resume(r.horsCible)).toEqual([]);
  });
});

test.describe("Relevé — autres chemins qui tapent : « Reprendre » et « Lancer le diagnostic »", () => {
  test("A27 « Reprendre » après une pause pendant laquelle l'utilisateur a commencé un message (curseur dans son objet) : rien n'est tapé dans ce message", async ({ page, worker }) => {
    // BUG : même cause que A05/A06 (adresses/panel.js:476-479) : à la reprise, O.champA() = « À » du message de l'utilisateur.
    await ouvrir(page, worker);
    await page.evaluate(() => (window as any).ow.afficherMessage(null));
    await demarrerReleve(page);
    await attendreRecherches(page, 2);
    await arreterReleve(page);
    // L'utilisateur écrit un message pendant la pause.
    await page.locator("#nouveau-courrier").click();
    const id = (await etat(page)).actif!;
    await page.locator(`[data-zone="a-${id}"]`).click(); await page.keyboard.type("paul.durand@exemple.fr;", { delay: 5 });
    await page.locator(`[data-zone="objet-${id}"]`).click(); await page.keyboard.type("Planning");
    const avant = (await etat(page)).redactions.find((x) => x.id === id)!;
    await viderJournal(page);
    await ao(page).locator("#ao-run").click(); // « Reprendre »
    await attendreDebutOuRefus(page, 20000);
    await page.waitForTimeout(2500);
    if (await releveEnCours(page)) await arreterReleve(page);
    const apres = (await etat(page)).redactions.find((x) => x.id === id)!;
    expect((await saisiesExtension(page)).filter((x) => x.zone.endsWith("-" + id)).map((x) => x.zone + ":" + x.inputType)).toEqual([]);
    expect({ a: apres.a, aTexte: apres.aTexte, objet: apres.objet }).toEqual({ a: avant.a, aTexte: avant.aTexte, objet: avant.objet });
  });

  test("A28 ⚙ « Lancer le diagnostic » avec le curseur dans la réponse de l'utilisateur : ni « a » tapé ni lettres effacées dans son texte", async ({ page, worker }) => {
    // BUG : adresses/panel.js:508-528 le diagnostic arme le champ du curseur sans condition (armement.js:132-145
    // pickHere), tape « a » puis efface 3 caractères (panel.js:526) : « a » + 2 lettres de l'utilisateur disparaissent.
    await ouvrir(page, worker);
    const id = await reponseUtilisateur(page);
    await cliquerFinDe(page, `[data-zone="corps-${id}"] .elementToProof`);
    const avant = (await etat(page)).redactions.find((x) => x.id === id)!;
    await ao(page).locator("#ao-gear").click();
    const dl = page.waitForEvent("download", { timeout: 20000 }).catch(() => null);
    await ao(page).locator("#ao-diag").click();
    await dl;
    await page.waitForTimeout(500);
    const apres = (await etat(page)).redactions.find((x) => x.id === id)!;
    expect((await saisiesExtension(page)).map((x) => x.zone + ":" + x.inputType)).toEqual([]);
    expect(apres.corps).toBe(avant.corps);
  });

  test("A29 ⚙ « Lancer le diagnostic », curseur sur la liste des messages après avoir écrit dans Teams : rien n'est tapé dans Teams", async ({ page, worker }) => {
    // BUG : armement.js:135 pickHere retombe sur lastEditable (dernier champ éditable focalisé, quel qu'il soit) :
    // le diagnostic tape « a » et efface 3 caractères dans le brouillon Teams.
    await ouvrir(page, worker);
    await page.locator("#btn-teams").click();
    await page.locator("#teams-saisie").click(); await page.keyboard.type("Je te rappelle demain", { delay: 5 });
    await page.locator('#liste [role="option"]').nth(2).click(); // l'utilisateur lit un autre message
    const avant = (await etat(page)).teams;
    await ao(page).locator("#ao-gear").click();
    const dl = page.waitForEvent("download", { timeout: 20000 }).catch(() => null);
    await ao(page).locator("#ao-diag").click();
    await dl;
    await page.waitForTimeout(500);
    expect((await saisiesExtension(page)).filter((x) => x.zone === "teams").map((x) => x.inputType + ":" + (x.data ?? ""))).toEqual([]);
    expect((await etat(page)).teams).toBe(avant);
  });
});
