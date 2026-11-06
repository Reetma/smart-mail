// Audit « Intégrité de chaque envoi » : reproductions CIBLÉES et déterministes des
// divergences trouvées par l'exploration à graines (audit-integrite-envoi-fuzz.spec.ts),
// plus des cas sûrs qui doivent passer (non-régression).
// Chaque test « BUG » exprime le comportement sûr attendu et ÉCHOUE aujourd'hui.
//   node node_modules/@playwright/test/cli.js test tests/audit-integrite-envoi-fuzz-cibles.spec.ts --workers=2 --reporter=line
import { test, expect, executer, rapport } from "./audit/integrite-envoi-fuzz/banc";
import { verifier, SIGNATURES, type Plan, type Releve } from "./audit/integrite-envoi-fuzz/generateur";
import { cfgBase, planBase } from "./audit/integrite-envoi-fuzz/cas";

test.describe.configure({ mode: "parallel" });

function sansDivergence(plan: Plan, r: Releve): void {
  const d = verifier(plan, r);
  if (process.env.FUZZ_TRACE) console.log(`\n### ${test.info().title}\n` + rapport(plan, r, d));
  expect(d, rapport(plan, r, d)).toEqual([]);
}
// Rapport complet (ce qui est parti, journal) joint à chaque assertion ciblée.
const preuve = (plan: Plan, r: Releve) => rapport(plan, r, verifier(plan, r));
const servis = (r: Releve) => r.envoyes.flatMap((m) => [...m.a, ...m.cc, ...m.cci].map((x) => x.toLowerCase()));
const clicsEnvoyer = (r: Releve) => r.journal.filter((j) => j.type === "clic-envoyer").length;

// ------------------------------------------------------------ défauts (échouent aujourd'hui)

test("BUG E1 accents : « andré.martin@ » et « zoë.nguyen@ » ne partent jamais vers « martin@ » / « nguyen@ »", async ({ banc }) => {
  // BUG : moteur.js:60-68 (extraireAdresses, motif TROUVE en ASCII seul) coupe la partie
  // locale au premier caractère accentué : « andré.martin@exemple.fr » devient
  // « martin@exemple.fr », « zoë.nguyen@exemple.fr » devient « nguyen@exemple.fr » : deux
  // AUTRES personnes reçoivent le mail. « josé@exemple.fr » disparaît sans un mot.
  const plan = planBase({ destinataires: ["andré.martin@exemple.fr", "zoë.nguyen@exemple.fr", "josé@exemple.fr", "camille.leroy@exemple.fr"] });
  const r = await executer(banc, cfgBase(), plan);
  expect(servis(r), preuve(plan, r)).not.toContain("martin@exemple.fr");
  expect(servis(r), preuve(plan, r)).not.toContain("nguyen@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E1 suggestion voisine dans « À » : la pastille montre « jean.dupont@ » au lieu de « jean@ » → rien ne part", async ({ banc }) => {
  // BUG : outlook.js:367-372 (aValidee) et 396-406 (remplirA) acceptent N'IMPORTE quelle
  // pastille nouvelle : son identité n'est jamais comparée à l'adresse collée, même quand
  // la pastille affiche une autre adresse. VARIANTE plausible (non confirmée sur le vrai
  // Outlook) : la résolution remplace l'adresse tapée par la suggestion voisine.
  const cfg = cfgBase({ affichage: "adresse", remplacement: { proba: 1, quand: "collage" } });
  const plan = planBase({ destinataires: ["jean@exemple.fr"] });
  const r = await executer(banc, cfg, plan);
  expect(servis(r), preuve(plan, r)).not.toContain("jean.dupont@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E1 suggestion voisine dans « Cci » (adresse en info-bulle) : aucun destinataire remplacé", async ({ banc }) => {
  // BUG : outlook.js:452-458 (remplirCci) valide au NOMBRE de pastilles : une pastille
  // « MARTIN Alice » (info-bulle alice.martin@) pour « alice@ » compte comme bonne.
  const cfg = cfgBase({ affichage: "nom+titre", remplacement: { proba: 1, quand: "collage" } });
  const plan = planBase({ destinataires: ["camille.leroy@exemple.fr", "alice@exemple.fr", "hugo.blanc@exemple.org"], cci: true, paquet: 3 });
  const r = await executer(banc, cfg, plan);
  expect(servis(r), preuve(plan, r)).not.toContain("alice.martin@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E3 corps tronqué au collage (dernier paragraphe perdu) : rien ne part", async ({ banc }) => {
  // BUG : outlook.js:558-565 (remplirCorps) ne vérifie que l'EXCÈS de texte, les 30
  // premiers caractères de la première ligne et la présence des liens : un corps amputé
  // de sa fin part tel quel. VARIANTE plausible : l'éditeur ne garde qu'une partie du collage.
  const cfg = cfgBase({ filtreCorps: { styles: true, couper: "dernier-bloc" } });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG E3 corps coupé à 60 % (texte d'un lien compris) : rien ne part", async ({ banc }) => {
  // BUG : même cause (outlook.js:558-565) : le lien est présent (même adresse) mais son
  // texte et toute la suite du message manquent.
  const cfg = cfgBase({ filtreCorps: { styles: true, couper: "apres-n" } });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG E3 reste d'un message précédent dans le corps neuf : rien ne part", async ({ banc }) => {
  // BUG : outlook.js:551 (avantTexte) prend TOUT le texte déjà présent pour « la
  // signature » : un reste (« Bonjour Paul, … devis ») part avec le message. VARIANTE
  // plausible : rédaction neuve qui garde du texte (brouillon restauré, éditeur réutilisé).
  const cfg = cfgBase({ restes: [{ ou: "corps", html: "Bonjour Paul, suite à notre échange de ce matin, je te renvoie le devis.", position: "debut", redaction: 1 }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.some((m) => m.corpsTexte.includes("devis")), preuve(plan, r)).toBe(false);
  sansDivergence(plan, r);
});

test("BUG E1 rédaction qui garde un Cc : envoi « À » sans destinataire en trop", async ({ banc }) => {
  // BUG : le champ Cc n'est JAMAIS contrôlé (outlook.js:149 champCc ne sert qu'à délimiter
  // les lignes) : un Cc déjà rempli part avec chaque mail.
  const cfg = cfgBase({ restes: [{ ou: "cc", adresse: "paul.martin@exemple.fr", redaction: 1 }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(servis(r), preuve(plan, r)).not.toContain("paul.martin@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E1 rédaction qui garde un Cci : envoi « À » sans destinataire caché en trop", async ({ banc }) => {
  // BUG : en mode « À », outlook.js:386-409 (remplirA) ne regarde pas le Cci.
  const cfg = cfgBase({ restes: [{ ou: "cci", adresse: "paul.martin@exemple.fr", redaction: 2 }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(servis(r), preuve(plan, r)).not.toContain("paul.martin@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E1 rédaction qui garde un Cc, envoi en Cci : le paquet ne part pas avec un Cc", async ({ banc }) => {
  // BUG : outlook.js:471 (remplirCci) vérifie « À » vide mais pas Cc.
  const cfg = cfgBase({ restes: [{ ou: "cc", adresse: "paul.martin@exemple.fr", redaction: 1 }] });
  const plan = planBase({ destinataires: ["camille.leroy@exemple.fr", "hugo.blanc@exemple.org", "lea.roux@exemple.fr"], cci: true, paquet: 3 });
  const r = await executer(banc, cfg, plan);
  expect(servis(r), preuve(plan, r)).not.toContain("paul.martin@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E3 complément qui réécrit le corps juste avant « Envoyer » : le mail ne part pas sans le message", async ({ banc }) => {
  // BUG : moteur.js:304-313 + outlook.js:719-728 : AUCUN contrôle final entre la fin du
  // remplissage et le clic « Envoyer ». Un complément de signature (body.setAsync sur un
  // corps lu à l'ouverture, VARIANTE plausible) efface le message : les mails partent
  // avec la seule signature. L'instant est choisi par l'horloge de test (étape « Envoyer »).
  const cfg = cfgBase({ tardifs: [{ quand: "etape:Envoyer", quoi: "corps-ecrase" }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.filter((m) => !m.corpsTexte.includes("Chers étudiants")), preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG E1 pastille en double apparue après la validation (sortie du champ « À ») : rien ne part", async ({ banc }) => {
  // BUG : outlook.js:406 (« Plusieurs destinataires ») n'est évalué qu'une fois, juste
  // après la validation ; une pastille ajoutée à la sortie du champ part aussi.
  const cfg = cfgBase({ tardifs: [{ quand: "blur:a", quoi: "doublon-tardif", champ: "a" }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.filter((m) => m.a.length > 1), preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG E1 résolution qui change la pastille après validation (« zoe@ » → « NGUYEN Zoé ») : rien ne part", async ({ banc }) => {
  // BUG : même absence de contrôle final : la pastille « zoe@exemple.fr » devient
  // « NGUYEN Zoé » (zoe.nguyen@) à la sortie du champ, le mail part vers une autre
  // personne et « zoe@ » est noté envoyé. VARIANTE plausible (re-résolution différée).
  const cfg = cfgBase({ tardifs: [{ quand: "blur:a", quoi: "remplacement-tardif", champ: "a" }] });
  const plan = planBase({ destinataires: ["zoe@exemple.fr"] });
  const r = await executer(banc, cfg, plan);
  expect(servis(r), preuve(plan, r)).not.toContain("zoe.nguyen@exemple.fr");
  sansDivergence(plan, r);
});

test("BUG E4 importance haute confirmée puis perdue avant « Envoyer » : rien ne part", async ({ banc }) => {
  // BUG : outlook.js:678-715 confirme l'importance une fois ; rien ne la revérifie avant
  // le clic (re-rendu du ruban qui la remet à « normale », VARIANTE plausible).
  const cfg = cfgBase({ tardifs: [{ quand: "etape:Envoyer", quoi: "importance-reset" }] });
  const plan = planBase({ importance: true });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.filter((m) => m.importance !== "haute"), preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG E3 lien du message retiré au collage mais présent dans la signature : rien ne part", async ({ banc }) => {
  // BUG : outlook.js:563-565 cherche chaque lien prévu dans TOUT le corps (signature
  // comprise) et sans compter : le lien de la signature (même site) masque la perte du
  // lien du message. VARIANTE plausible : l'éditeur défait les liens collés.
  const cfg = cfgBase({ signature: { html: SIGNATURES.lien!, position: "apres" }, filtreCorps: { styles: true, retirerLiens: 1 } });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG E6 rédaction imbriquée profondément : jamais de clic sur le « Envoyer » d'une conversation Teams", async ({ banc }) => {
  // BUG : outlook.js:154-163 (zoneRedaction) abandonne au-delà de 25 niveaux, puis
  // outlook.js:723 (envoyer) clique le PREMIER « Envoyer » visible de toute la page :
  // ici celui de la conversation Teams ouverte. Le destinataire est ensuite noté « à
  // vérifier » et rangé dans « Déjà envoyés » alors que rien n'est parti. (Profondeur réelle
  // d'Outlook inconnue : VARIANTE.)
  const cfg = cfgBase({ profondeurObjet: 24, autreEnvoyer: true });
  const plan = planBase({ destinataires: ["camille.leroy@exemple.fr"] });
  const r = await executer(banc, cfg, plan);
  expect(r.journal.filter((j) => j.type === "clic-interdit").map((j) => j.detail.role), preuve(plan, r)).not.toContain("interdit:teams-envoyer");
  sansDivergence(plan, r);
});

for (const cache of ["hors-ecran", "opacite", "hauteur0"] as const) {
  test(`BUG E6 rédaction masquée (${cache}) posée avant la vraie : rien n'y est tapé, rien n'en part`, async ({ banc }) => {
    // BUG : outlook.js:67-75 (visible) ne tient compte ni de la position hors écran, ni de
    // l'opacité, ni d'un parent de hauteur nulle : champA/champObjet/champCorps (146-150)
    // prennent la rédaction masquée, premier élément dans l'ordre du document. VARIANTE
    // plausible (rédaction pré-rendue ou repliée, forme réelle inconnue).
    const cfg = cfgBase({ doublonsCaches: [{ quoi: "redaction-vide", cache, ordre: "avant", quand: "avec-redaction" }] });
    const plan = planBase({ importance: true });
    const r = await executer(banc, cfg, plan);
    expect(r.journal.filter((j) => j.type === "saisie-redaction-cachee"), preuve(plan, r)).toHaveLength(0);
    sansDivergence(plan, r);
  });
}

test("BUG E6 vieux bouton « Envoyer » masqué (opacité 0) avant le vrai : jamais cliqué, destinataire jamais perdu", async ({ banc }) => {
  // BUG : outlook.js:67-75 (visible) + 723 (envoyer) : le premier « Envoyer » de la zone
  // dans l'ordre du document est un bouton périmé transparent ; il est cliqué, rien ne
  // part, et moteur.js:316/399 classe le destinataire « à vérifier » puis le range dans
  // « Déjà envoyés ». VARIANTE plausible (reste d'une rédaction précédente, animation).
  const cfg = cfgBase({ doublonsCaches: [{ quoi: "envoyer", cache: "opacite", ordre: "avant", quand: "avec-redaction" }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.journal.filter((j) => j.type === "clic-interdit").map((j) => j.detail.role), preuve(plan, r)).not.toContain("interdit:envoyer-perime");
  sansDivergence(plan, r);
});

test("BUG E3 signature en tête : le message est collé DANS le bloc Signature, que le complément réapplique avant « Envoyer »", async ({ banc }) => {
  // BUG : outlook.js:238-247 (curseurDebut) descend dans le premier élément du corps : si
  // c'est le bloc « Signature » (signature en tête, VARIANTE plausible), le message y est
  // collé ; un complément qui réapplique sa signature (setSignatureAsync remplace ce bloc)
  // l'efface, et faute de contrôle final (moteur.js:304-313) le mail part sans message.
  const cfg = cfgBase({ signature: { html: SIGNATURES.simple!, position: "avant" }, tardifs: [{ quand: "etape:Envoyer", quoi: "signature-remplace" }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.filter((m) => !m.corpsTexte.includes("Chers étudiants")), preuve(plan, r)).toHaveLength(0);
  sansDivergence(plan, r);
});

test("BUG perte : pastille refusée par Outlook à l'envoi → destinataire « à vérifier » et rangé dans « Déjà envoyés »", async ({ banc }) => {
  // BUG (mineur) : outlook.js:367-372 compte une pastille INVALIDE (soulignée en rouge,
  // aria-invalid) comme validée ; Outlook refuse l'envoi (« n'a pas pu être envoyé »),
  // moteur.js:316 classe le destinataire « à vérifier » et moteur.js:399 l'ajoute à
  // « Déjà envoyés » : il ne recevra jamais rien, alors que rien n'est parti. Attendu :
  // arrêt AVANT « Envoyer », destinataire laissé en attente.
  const cfg = cfgBase({ echec: { proba: 1, mode: "invalide" } });
  const plan = planBase({ destinataires: ["camille.leroy@exemple.fr"] });
  const r = await executer(banc, cfg, plan);
  expect(r.journal.filter((j) => j.type === "clic-interdit"), preuve(plan, r)).toHaveLength(0); // jamais « OK » sur l'alerte
  expect(clicsEnvoyer(r), preuve(plan, r)).toBe(0);
  expect(r.statuts["camille.leroy@exemple.fr"], preuve(plan, r)).not.toBe("verifier");
});

// ------------------------------------------------------------ cas sûrs (doivent passer)

test("sûr : résolution lente (texte affiché pendant 2,5 s), pastilles au nom, ids changeants, enveloppes", async ({ banc }) => {
  const cfg = cfgBase({ delai: { max: 2500, mode: "texte-reste" }, enveloppes: 3, idsChangeants: true });
  const plan = planBase({ destinataires: ["jean.dupont@exemple.fr", "Alice.Martin@EXEMPLE.FR", "o'brien+promo@exemple.ie"] });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(3);
  sansDivergence(plan, r);
});

test("sûr : résolution lente (texte qui disparaît puis pastille), Cci par 7 sur 20 adresses variées", async ({ banc }) => {
  const cfg = cfgBase({ delai: { max: 1500, mode: "texte-disparait" }, cciVisible: false, basculeCci: "aria" });
  const dests = Array.from({ length: 20 }, (_, i) => [`Etudiant${i}.Nom@EXEMPLE.FR`, `anne-laure.le-goff${i}@mail.departement-des-sciences-humaines-et-sociales.universite-de-lorraine.fr`, `n'diaye+ia${i}@x.io`, `jean_baptiste${i}@exemple.museum`][i % 4]!);
  const plan = planBase({ destinataires: dests, cci: true, paquet: 7 });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.map((m) => m.cci.length), preuve(plan, r)).toEqual([7, 7, 6]);
  sansDivergence(plan, r);
});

test("sûr : Cci par 50, champ Cci déjà affiché, pastilles à l'adresse", async ({ banc }) => {
  const cfg = cfgBase({ cciVisible: true, affichage: "adresse" });
  const dests = Array.from({ length: 50 }, (_, i) => `personne.${i}@ecole-${i % 3}.fr`);
  const plan = planBase({ destinataires: dests, cci: true, paquet: 50 });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes.map((m) => m.cci.length), preuve(plan, r)).toEqual([50]);
  sansDivergence(plan, r);
});

test("sûr : re-rendus React (en-tête après le collage, corps à l'entrée, objet à la sortie)", async ({ banc }) => {
  const cfg = cfgBase({ tardifs: [{ quand: "colle:a", quoi: "rerendu-entete" }, { quand: "focus:corps", quoi: "rerendu-corps" }, { quand: "blur:objet", quoi: "rerendu-objet" }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  sansDivergence(plan, r); // envoi correct, ou arrêt propre sans envoi
});

test("sûr : pastille en double dès le collage → arrêt, rien ne part (« À »)", async ({ banc }) => {
  const cfg = cfgBase({ doublon: { proba: 1, quand: "collage" } });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("rien n'est parti");
  sansDivergence(plan, r);
});

test("sûr : pastille en double dès le collage → arrêt, rien ne part (« Cci »)", async ({ banc }) => {
  const cfg = cfgBase({ doublon: { proba: 1, quand: "collage" } });
  const plan = planBase({ destinataires: ["a1@exemple.fr", "a2@exemple.fr"], cci: true, paquet: 2 });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("rien n'est parti");
  sansDivergence(plan, r);
});

test("sûr : résolution qui échoue (texte resté) → arrêt avec message, rien ne part", async ({ banc }) => {
  const cfg = cfgBase({ echec: { proba: 1, mode: "texte" } });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toMatch(/n'a pas été validée dans « À »/);
  sansDivergence(plan, r);
});

test("sûr : liens réécrits en « safelinks » au collage → arrêt « Lien absent », rien ne part", async ({ banc }) => {
  const cfg = cfgBase({ filtreCorps: { styles: true, safelinks: 1 } });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("Lien absent du mail");
  sansDivergence(plan, r);
});

test("sûr : collage refusé (rien n'entre) → arrêt, rien ne part", async ({ banc }) => {
  const cfg = cfgBase({ collerCorps: "bloque" });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("Le texte n'apparaît pas dans le corps du message");
  sansDivergence(plan, r);
});

test("sûr : collage ignoré par l'éditeur (repli de l'extension) et styles retirés → message intact", async ({ banc }) => {
  const cfg = cfgBase({ collerCorps: "ignore", collerDest: "ignore" });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(2);
  sansDivergence(plan, r);
});

test("sûr : objet de plus de 255 caractères (maxlength réel) → arrêt, rien ne part", async ({ banc }) => {
  const plan = planBase({ objet: "Programme détaillé de la formation ; ".repeat(9).trim() });
  const r = await executer(banc, cfgBase(), plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("L'objet collé ne correspond pas");
  sansDivergence(plan, r);
});

for (const quoi of ["redaction-vide", "envoyer", "a-reponse"] as const) {
  for (const cache of ["display", "visibility"] as const) {
    test(`sûr : doublon caché (${quoi}, ${cache}) ignoré, envoi correct`, async ({ banc }) => {
      const cfg = cfgBase({ doublonsCaches: [{ quoi, cache, ordre: "avant", quand: "avec-redaction" }] });
      const plan = planBase();
      const r = await executer(banc, cfg, plan);
      expect(r.envoyes, preuve(plan, r)).toHaveLength(2);
      sansDivergence(plan, r);
    });
  }
}

test("sûr : brouillon de l'utilisateur ouvert (rédaction visible) → refus de démarrer, rien n'y est tapé", async ({ banc }) => {
  const cfg = cfgBase({ doublonsCaches: [{ quoi: "redaction-brouillon", cache: "hors-ecran", ordre: "avant", quand: "toujours" }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("Un message est déjà ouvert en rédaction");
  sansDivergence(plan, r);
});

test("sûr : « À » déjà rempli par un reste → arrêt, rien ne part", async ({ banc }) => {
  const cfg = cfgBase({ restes: [{ ou: "a", adresse: "paul.martin@exemple.fr", redaction: 1 }] });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(0);
  expect(r.notice, preuve(plan, r)).toContain("Le champ « À » n'est pas vide");
  sansDivergence(plan, r);
});

test("sûr : fenêtre « Letsignit Add-in » passagère à l'ouverture → attente, puis envoi correct", async ({ banc }) => {
  const cfg = cfgBase({ fenetreOuverture: 3000 });
  const plan = planBase();
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(2);
  sansDivergence(plan, r);
});

test("sûr : signature AVANT la ligne du curseur et texte de signature retiré → message intact", async ({ banc }) => {
  const cfg = cfgBase({ signature: { html: SIGNATURES.ieseg!, position: "avant" } });
  const plan = planBase({ sansSignature: true });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(2);
  sansDivergence(plan, r);
});

test("sûr : signature posée par le complément juste après le collage, texte retiré au second passage", async ({ banc }) => {
  const cfg = cfgBase({ signature: null, tardifs: [{ quand: "colle:corps", quoi: "signature-fin" }] });
  const plan = planBase({ sansSignature: true });
  const r = await executer(banc, cfg, plan);
  expect(r.envoyes, preuve(plan, r)).toHaveLength(2);
  sansDivergence(plan, r);
});

test("sûr : suggestion voisine mais pastille au NOM seul — limite connue, aucun moyen de la voir (journalisé)", async ({ banc }) => {
  // Pas un défaut de l'extension : la page ne montre nulle part l'adresse retenue. Ce test
  // documente le risque résiduel ; il vérifie seulement qu'aucune autre règle ne casse.
  const cfg = cfgBase({ affichage: "nom", remplacement: { proba: 1, quand: "collage" } });
  const plan = planBase({ destinataires: ["jean@exemple.fr"] });
  const r = await executer(banc, cfg, plan);
  test.info().annotations.push({ type: "risque résiduel", description: `envoyé à : ${servis(r).join(", ") || "personne"}` });
  expect(r.journal.filter((j) => j.type === "clic-interdit"), preuve(plan, r)).toHaveLength(0);
});
