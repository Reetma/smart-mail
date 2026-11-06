// Générateur à graines : pour chaque graine, une fausse page Outlook (cfg, lue par
// fausse-outlook-fuzz.js) et un plan d'envoi (destinataires, objet, message,
// réglages). Puis l'ORACLE : ce que la fausse page a « envoyé » comparé au prévu.
// Tout est déterministe : même graine, mêmes variantes, même plan.

export function mulberry32(graine: number): () => number {
  let a = graine >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export type Contact = { nom: string; adresse: string };
export type Tardif = { quand: string; quoi: string; champ?: string; redaction?: number; proba?: number };
export type DoublonCache = { quoi: "redaction-brouillon" | "redaction-vide" | "envoyer" | "a-reponse"; cache: "display" | "visibility" | "hors-ecran" | "opacite" | "hauteur0"; ordre: "avant" | "apres"; quand: "toujours" | "avec-redaction" };
export type Cfg = {
  graine: number;
  nouveauSansAria?: boolean;
  enveloppes: number;
  profondeurRacine: number;
  profondeurObjet: number;
  idsChangeants: boolean;
  affichage: "nom" | "adresse" | "nom+titre";
  collerDest: "pastilles" | "ignore";
  delai: null | { max: number; mode: "texte-reste" | "texte-disparait" };
  remplacement: null | { proba: number; quand: "collage" };
  echec: null | { proba: number; mode: "texte" | "invalide" };
  doublon: null | { proba: number; quand: "collage" };
  ccVisible: boolean;
  cciVisible: boolean;
  basculeCci: "texte" | "aria" | "aria-afficher";
  collerCorps: "html" | "ignore" | "bloque" | "texte";
  filtreCorps: { styles?: boolean; safelinks?: number; retirerLiens?: number; couper?: null | "dernier-bloc" | "listes" | "apres-n" };
  signature: null | { html: string; position: "apres" | "avant" };
  signatureComplement?: string;
  tardifs: Tardif[];
  restes: Array<{ ou: "corps" | "cc" | "cci" | "a"; html?: string; adresse?: string; position?: "debut" | "avant-signature"; redaction?: number }>;
  doublonsCaches: DoublonCache[];
  autreEnvoyer: boolean;
  rerenduFocus: boolean;
  fenetreOuverture: number;
  annuaire: Contact[];
  listeMessages: Array<{ nom: string; adresse: string; objet: string }>;
};
export type Plan = {
  graine: number;
  destinataires: string[]; // ce que l'utilisateur VEUT servir (adresses telles qu'il les a écrites)
  saisie: string; // contenu collé dans « Destinataires » du panneau
  objet: string;
  html: string; // message collé dans le panneau
  cci: boolean;
  paquet: number;
  importance: boolean;
  sansSignature: boolean;
  resume: string[]; // variantes tirées (pour lire un échec)
};

// ------------------------------------------------------------ matériaux
const PRENOMS = ["jean", "marie", "lucas", "paul-henri", "anne-laure", "zoe", "hugo", "ines", "noah", "lea", "mohamed", "yasmine", "thomas", "camille", "jean-baptiste", "marie-claire", "kevin", "sarah"];
const NOMS = ["dupont", "martin", "lefevre", "nguyen", "o'brien", "d'artagnan", "saint-exupery", "garcia", "muller", "rousseau", "benali", "kowalski", "le-goff", "n'diaye"];
const DOMAINES = [
  "exemple.fr", "ieseg.fr", "etu.univ-paris1.fr", "reetma.com", "exemple.com", "x.io", "exemple.museum",
  "edu.ecole-superieure-de-commerce-international.fr",
  "mail.departement-des-sciences-humaines-et-sociales.universite-de-lorraine.fr",
  "sub1.sub2.sub3.exemple-organisation-internationale.org",
];
// Parties locales accentuées (IDN) : écrites ainsi par l'utilisateur.
const IDN_LOCAUX = ["élodie.durand", "andré.martin", "zoë.nguyen", "françois.dupont", "josé", "hélène-marie.lefèvre"];
const IDN_DOMAINES = ["société-générale.fr", "bücher.de", "école-polytechnique.fr"];
const PHRASES = [
  "Chers étudiants,", "Bonjour à toutes et à tous,", "Les inscriptions à la certification IA ferment le 31 octobre.",
  "Vous pourrez ajouter votre certificat de réussite à votre profil LinkedIn et à votre CV.",
  "71 % des dirigeants préféreraient recruter une personne formée à l’IA.", "Les places sont limitées : réservez la vôtre dès aujourd’hui.",
  "Le programme dure six semaines, à raison de deux heures par semaine.", "N’hésitez pas à répondre à ce message pour toute question.",
  "Rendez-vous le 27 octobre à Paris, salle « Lumière ».", "Cordialement,", "L’équipe pédagogique de Reetma School",
  "Merci de confirmer votre présence avant vendredi 18 h.", "Cette formation est gratuite pour les étudiants de première année.",
];
const LIENS = [
  "https://reetma.com/ecole", "https://exemple.fr/inscription?promo=2026&utm_source=mail", "https://www.ieseg.fr/programmes/",
  "mailto:contact@exemple.fr", "https://exemple.org/doc#section-2", "https://Exemple.fr/Chemin/Majuscules",
];
const TEXTES_LIENS = ["ici", "OBTENEZ VOTRE CERTIFICATION AETHER EN IA", "le programme", "écrivez-nous", "la documentation", null];
const OBJETS = [
  "Certification IA - Avant 31 Oct.", "Invitation : journée portes ouvertes du 27/10", "Votre inscription à l’atelier « IA générative »",
  "Rappel — dernier délai lundi 9 h", "Réunion d'information (salle B-204)", "🎓 Certification Reetma : il reste 12 places",
  "Objet très long : " + "programme détaillé de la formation, horaires, lieux et modalités d’inscription ; ".repeat(3).trim(),
];
const IMG = "data:image/gif;base64,R0lGODlhAQABAAAAACw=";
export const SIGNATURES: Record<string, string> = {
  simple: "<div>--</div><div>Paul DURANDEAU</div><div>Reetma School · Paris</div>",
  ieseg:
    `<table><tbody><tr><td><a href="https://facebook.com/ieseg"><img alt="facebook" src="${IMG}"></a><a href="https://linkedin.com/school/ieseg"><img alt="linkedin" src="${IMG}"></a></td>` +
    `<td style="border-left:2px solid #000"><div>DURANDEAU Paul</div><div>-BA1</div><div>Paris</div><img alt="IESEG" src="${IMG}"></td></tr></tbody></table>`,
  // même lien que le message (site de l'école) : plausible pour une signature d'établissement
  lien: '<div>Paul DURANDEAU — Reetma School</div><div><a href="https://reetma.com/ecole">reetma.com/ecole</a></div>',
};

function cap(s: string) { return s.split(/([.\-'])/).map((x) => (x.length > 1 ? x[0]!.toUpperCase() + x.slice(1) : x)).join(""); }

// ------------------------------------------------------------ plan + page
export function genererGraine(graine: number): { cfg: Cfg; plan: Plan } {
  const r = mulberry32(graine * 2654435761);
  const p = (x: number) => r() < x;
  const un = <T,>(t: readonly T[]): T => t[Math.floor(r() * t.length)]!;
  const entre = (a: number, b: number) => a + Math.floor(r() * (b - a + 1));
  const resume: string[] = [];

  // --- destinataires
  const cci = p(0.4);
  const paquet = cci ? un([1, 2, 3, 5, 7, 10, 20, 30, 50]) : 1;
  const nbMails = cci ? entre(1, 3) : entre(1, 4);
  const total = cci ? Math.max(1, Math.min(60, paquet * nbMails - (nbMails > 1 ? entre(0, Math.max(0, paquet - 1)) : 0))) : nbMails;
  const annuaire: Contact[] = [];
  const vus = new Set<string>();
  const dests: string[] = [];
  const idn = p(0.05);
  let essais = 0;
  while (dests.length < total && essais++ < 1000) {
    const pr = un(PRENOMS), no = un(NOMS), dom = un(DOMAINES);
    const forme = un(["p.n", "p.n", "P.N", "p+tag", "pn42", "p_n", "court", "court", "long", "apostrophe", "tiret"]);
    let a = "";
    if (forme === "p.n") a = `${pr}.${no}@${dom}`;
    else if (forme === "P.N") a = `${cap(pr)}.${cap(no)}@${dom.toUpperCase()}`;
    else if (forme === "p+tag") a = `${pr}+${un(["promo", "news-2026", "ia"])}@${dom}`;
    else if (forme === "pn42") a = `${pr[0]}${no.replace(/'/g, "")}${entre(1, 99)}@${dom}`;
    else if (forme === "p_n") a = `${pr}_${no}@${dom}`;
    else if (forme === "court") a = `${pr}@${dom}`; // a souvent une « suggestion voisine » pr.nom@dom dans l'annuaire
    else if (forme === "long") a = `${pr}.${no}@${un(DOMAINES.slice(-3))}`;
    else if (forme === "apostrophe") a = `${pr}.${un(["o'brien", "d'artagnan", "n'diaye"])}@${dom}`;
    else a = `${un(["anne-laure", "jean-baptiste", "paul-henri"])}.${un(["saint-exupery", "le-goff"])}@${dom}`;
    if (vus.has(a.toLowerCase())) continue;
    vus.add(a.toLowerCase());
    dests.push(a);
    // annuaire : la moitié des destinataires y sont (la pastille montre alors le NOM)
    if (p(0.5)) annuaire.push({ nom: `${no.toUpperCase()} ${cap(pr)}`, adresse: a.toLowerCase().replace(/@(.*)$/, (m) => m.toLowerCase()) });
    // suggestion voisine pour les formes courtes
    if (forme === "court") annuaire.push({ nom: `${un(NOMS).toUpperCase()} ${cap(pr)}`, adresse: `${pr}.${un(NOMS).replace(/'/g, "")}@${dom}` });
  }
  if (idn && dests.length) {
    const i = Math.floor(r() * dests.length);
    const x = p(0.8) ? `${un(IDN_LOCAUX)}@${un(DOMAINES.slice(0, 5))}` : `contact@${un(IDN_DOMAINES)}`;
    if (!vus.has(x.toLowerCase())) { dests[i] = x; resume.push("IDN:" + x); }
  }
  for (let i = 0; i < 6; i++) annuaire.push({ nom: `${un(NOMS).toUpperCase()} ${cap(un(PRENOMS))}`, adresse: `${un(PRENOMS)}.${un(NOMS).replace(/'/g, "")}${i}@${un(DOMAINES)}` });

  const format = un(["lignes", "lignes", "lignes", "points-virgules", "virgules", "noms"]);
  const saisie = format === "lignes" ? dests.join("\n")
    : format === "points-virgules" ? dests.join("; ")
    : format === "virgules" ? dests.join(", ")
    : dests.map((a, i) => `${cap(PRENOMS[i % PRENOMS.length]!)} ${NOMS[i % NOMS.length]!.toUpperCase()} <${a}>`).join("\n");
  resume.push(`mode:${cci ? "cci×" + paquet : "À"}`, `dest:${dests.length}`, `format:${format}`);

  // --- message
  const nbBlocs = entre(2, 6);
  const liensPris: string[] = [];
  const blocs: string[] = [];
  for (let i = 0; i < nbBlocs; i++) {
    let t = un(PHRASES);
    if (p(0.3)) t = t.replace(/(\S+)$/, "<strong>$1</strong>");
    if (p(0.15)) t = t.replace(/^(\S+)/, "<em>$1</em>");
    if (p(0.45) && liensPris.length < 3) {
      const href = un(LIENS), txt = un(TEXTES_LIENS) ?? href.replace(/^mailto:/, "");
      liensPris.push(href);
      t += ` → <a href="${href.replace(/&/g, "&amp;")}">${txt.replace(/&/g, "&amp;")}</a>`;
    }
    blocs.push(`<div>${t}</div>`);
    if (p(0.5)) blocs.push("<div><br></div>");
  }
  if (p(0.2)) blocs.splice(1, 0, `<ul><li>${un(PHRASES)}</li><li>${un(PHRASES)}</li></ul>`);
  const html = blocs.join("");
  const objet = p(0.05) ? OBJETS[OBJETS.length - 1]! : un(OBJETS.slice(0, -1)); // l'objet de plus de 255 caractères reste rare

  // --- page
  const sigType = un(["aucune", "simple", "simple", "ieseg", "ieseg", "lien"]);
  const sigPos = un(["apres", "apres", "apres", "apres", "avant"]) as "apres" | "avant";
  const sigTardive = sigType !== "aucune" && p(0.15);
  // Pastille au nom seul : une suggestion voisine y serait indétectable (rien dans la
  // page ne donne l'adresse) ; on ne la tire que si la pastille montre l'adresse.
  const affichage = un(["nom", "nom", "adresse", "nom+titre"] as const);
  const cfg: Cfg = {
    graine,
    nouveauSansAria: p(0.3),
    enveloppes: un([0, 0, 1, 2, 3]),
    profondeurRacine: un([0, 0, 0, 2, 5]),
    profondeurObjet: p(0.08) ? entre(18, 24) : entre(0, 4),
    idsChangeants: p(0.5),
    affichage,
    collerDest: p(0.25) ? "ignore" : "pastilles",
    delai: p(0.45) ? { max: un([300, 800, 1500, 2500]), mode: un(["texte-reste", "texte-disparait"]) } : null,
    remplacement: p(0.12) && affichage !== "nom" ? { proba: 0.6, quand: "collage" } : null,
    echec: p(0.08) ? { proba: 0.35, mode: un(["texte", "invalide"]) } : null,
    doublon: p(0.08) ? { proba: 0.35, quand: "collage" } : null,
    ccVisible: p(0.1),
    cciVisible: p(0.25),
    basculeCci: un(["texte", "aria", "aria-afficher"]),
    collerCorps: un(["html", "html", "html", "html", "html", "html", "ignore", "ignore", "texte", "bloque"]),
    filtreCorps: {
      styles: p(0.5),
      safelinks: p(0.07) ? 0.8 : 0,
      retirerLiens: p(0.05) ? 0.8 : 0,
      couper: p(0.12) ? un(["dernier-bloc", "listes", "apres-n"]) : null,
    },
    signature: sigType === "aucune" || sigTardive ? null : { html: SIGNATURES[sigType]!, position: sigPos },
    signatureComplement: SIGNATURES[sigType === "aucune" ? "simple" : sigType],
    tardifs: [],
    restes: [],
    doublonsCaches: [],
    autreEnvoyer: p(0.1),
    rerenduFocus: p(0.5),
    fenetreOuverture: p(0.25) ? un([300, 1500, 4000]) : 0,
    annuaire,
    listeMessages: [
      { nom: "Paul MARTIN", adresse: "paul.martin@exemple.fr", objet: "Compte rendu" },
      ...dests.slice(0, 2).map((a) => ({ nom: a.split("@")[0]!, adresse: a.toLowerCase(), objet: "Question sur l'inscription" })),
    ],
  };
  if (sigTardive) {
    const ou = un(["signature-fin", "signature-fin", "signature-debut", "signature-ligne-vide", "corps-ecrase"]);
    const quand = un(["ouverture", "etape:Objet", "etape:Message", "colle:corps", "etape:Envoyer"]);
    cfg.tardifs.push({ quand, quoi: ou });
    resume.push(`signature-tardive:${ou}@${quand}`);
  }
  if (p(0.25)) {
    const quoi = un(["rerendu-entete", "rerendu-objet", "rerendu-corps", "rerendu-tout"]);
    const quand = un(["colle:a", "colle:cci", "blur:a", "blur:cci", "focus:objet", "focus:corps", "colle:corps", "etape:Objet", "etape:Message", "etape:Envoyer"]);
    cfg.tardifs.push({ quand, quoi });
    resume.push(`rerendu:${quoi}@${quand}`);
  }
  if (p(0.1)) {
    const t = un<Tardif>([
      { quand: "blur:a", quoi: "doublon-tardif", champ: "a" },
      { quand: "blur:cci", quoi: "doublon-tardif", champ: "cci" },
      { quand: "blur:a", quoi: "remplacement-tardif", champ: "a" },
      { quand: "blur:cci", quoi: "remplacement-tardif", champ: "cci" },
      { quand: "etape:Envoyer", quoi: "importance-reset" },
      { quand: "etape:Envoyer", quoi: "corps-ecrase" },
      { quand: "etape:Importance", quoi: "corps-ecrase" },
    ]);
    cfg.tardifs.push(t);
    resume.push(`tardif:${t.quoi}@${t.quand}`);
  }
  if (p(0.08)) {
    const ou = un(["corps", "corps", "cc", "cci"] as const);
    const reste = ou === "corps"
      ? { ou, html: un(["Bonjour Paul, suite à notre échange de ce matin, je te renvoie le devis.", "Merci pour votre retour.", PHRASES[2]!]), position: un(["debut", "avant-signature"] as const), redaction: entre(1, 2) }
      : { ou, adresse: "paul.martin@exemple.fr", redaction: entre(1, 2) };
    cfg.restes.push(reste);
    resume.push(`reste:${ou}@redaction${reste.redaction}`);
  }
  if (p(0.15)) {
    const d: DoublonCache = {
      quoi: un(["redaction-brouillon", "redaction-vide", "envoyer", "a-reponse"] as const),
      cache: un(["display", "visibility", "hors-ecran", "opacite", "hauteur0"] as const),
      ordre: un(["avant", "apres"] as const),
      quand: un(["toujours", "avec-redaction"] as const),
    };
    if (d.quoi === "envoyer") d.quand = "avec-redaction";
    cfg.doublonsCaches.push(d);
    resume.push(`doublon-cache:${d.quoi}/${d.cache}/${d.ordre}/${d.quand}`);
  }
  if (cfg.remplacement) resume.push("remplacement-voisin");
  if (cfg.echec) resume.push("echec:" + cfg.echec.mode);
  if (cfg.doublon) resume.push("doublon-pastille");
  if (cfg.delai) resume.push(`delai:${cfg.delai.max}/${cfg.delai.mode}`);
  if (cfg.collerDest === "ignore") resume.push("collage-dest-ignore");
  if (cfg.collerCorps !== "html") resume.push("collage-corps:" + cfg.collerCorps);
  if (cfg.filtreCorps.couper) resume.push("coupe:" + cfg.filtreCorps.couper);
  if (cfg.filtreCorps.safelinks) resume.push("safelinks");
  if (cfg.filtreCorps.retirerLiens) resume.push("liens-retires");
  if (cfg.profondeurObjet > 10) resume.push("profond:" + cfg.profondeurObjet);
  if (cfg.autreEnvoyer) resume.push("autre-envoyer-teams");
  resume.push(`affichage:${cfg.affichage}`, `signature:${sigType}${sigTardive ? "(tardive)" : "/" + sigPos}`);

  const plan: Plan = { graine, destinataires: dests, saisie, objet, html, cci, paquet, importance: p(0.3), sansSignature: p(0.25), resume };
  if (plan.importance) resume.push("importance");
  if (plan.sansSignature) resume.push("sans-texte-signature");
  // Variante « suggestion voisine » active : au moins une adresse courte (« prenom@dom »)
  // qui a une voisine « prenom.nom@dom » dans l'annuaire (tirages en fin : les autres
  // variantes de la graine ne bougent pas).
  const voisinable = (a: string) => annuaire.some((c) => { const [l, d2] = c.adresse.split("@"); const [l0, d0] = a.toLowerCase().split("@"); return d2 === d0 && l !== l0 && (l!.startsWith(l0 + ".") || l!.startsWith(l0 + "-")); });
  if ((cfg.remplacement || cfg.tardifs.some((t) => t.quoi === "remplacement-tardif")) && !dests.some(voisinable) && dests.length) {
    const k = Math.floor(r() * dests.length), pr = un(PRENOMS), dom = un(DOMAINES.slice(0, 6)), no = un(NOMS).replace(/'/g, "");
    const court = `${pr}@${dom}`;
    if (!vus.has(court)) {
      dests[k] = court;
      annuaire.push({ nom: `${no.toUpperCase()} ${cap(pr)}`, adresse: `${pr}.${no}@${dom}` });
      plan.saisie = format === "lignes" ? dests.join("\n") : format === "points-virgules" ? dests.join("; ") : format === "virgules" ? dests.join(", ") : dests.map((a, i) => `${cap(PRENOMS[i % PRENOMS.length]!)} ${NOMS[i % NOMS.length]!.toUpperCase()} <${a}>`).join("\n");
      resume.push("adresse-courte:" + court);
    }
  }
  return { cfg, plan };
}

// ------------------------------------------------------------ oracle
export type Envoye = {
  n: number; cachee: string | null; utilisateur: boolean; a: string[]; cc: string[]; cci: string[];
  objet: string; corpsTexte: string; liens: Array<{ href: string; texte: string }>; importance: string;
  signaturesTexte: string[]; signaturesLiens: string[]; restesTexte: string[];
};
export type Releve = {
  envoyes: Envoye[];
  journal: Array<{ type: string; detail: any }>;
  notice: string;
  noticeType: string;
  statuts: Record<string, string>; // adresse (minuscules) -> statut dans la campagne
  liste: string[]; // liste de l'envoi (pp-liste)
  prevu: { texte: string; liens: Array<{ href: string; texte: string }> } | null; // message de la campagne
  objetCampagne: string | null;
  fini: boolean;
  infos: string; // textes d'information du panneau (adresses refusées, compte…), relevés avant l'envoi
};

export const T = (s: string) => String(s || "").normalize("NFC").replace(/[^\p{L}\p{N}]/gu, "");
const sansBarre = (h: string) => String(h || "").replace(/\/+$/, "");
const minus = (a: string) => a.toLowerCase();

function permutations(t: string[]): string[] {
  if (t.length <= 1) return [t.join("")];
  const out: string[] = [];
  t.forEach((x, i) => permutations(t.filter((_, j) => j !== i)).forEach((y) => out.push(x + y)));
  return out;
}

// Divergences entre le prévu et le « envoyé ». Vide = intègre (ou arrêt propre).
export function verifier(plan: Plan, r: Releve): string[] {
  const d: string[] = [];
  const prevus = new Map(plan.destinataires.map((a) => [minus(a), a]));
  const servis = new Map<string, number>();
  if (!r.fini) d.push("BLOCAGE : l'envoi ne s'est pas terminé (ni fin, ni arrêt affiché)");

  r.envoyes.forEach((m, k) => {
    const tag = `mail ${k + 1}`;
    if (m.cachee || m.utilisateur) d.push(`E6 ${tag} : envoyé depuis une rédaction cachée/brouillon de l'utilisateur (${m.cachee || "brouillon"})`);
    const tous = [...m.a, ...m.cc, ...m.cci];
    tous.forEach((a) => { if (!prevus.has(minus(a))) d.push(`E1 ${tag} : destinataire NON PRÉVU « ${a} »`); });
    new Set(tous.map(minus)).forEach((a) => servis.set(a, (servis.get(a) || 0) + 1)); // un mail par adresse, même listée deux fois
    if (new Set(tous.map(minus)).size !== tous.length) d.push(`E1 ${tag} : même destinataire deux fois dans le mail (${tous.join(", ")})`);
    if (m.cc.length) d.push(`E1 ${tag} : Cc non vide (${m.cc.join(", ")})`);
    if (plan.cci) {
      if (m.a.length) d.push(`E1 ${tag} : « À » non vide en Cci (${m.a.join(", ")})`);
      if (m.cci.length > plan.paquet) d.push(`E1 ${tag} : ${m.cci.length} destinataires en Cci pour un paquet de ${plan.paquet}`);
    } else {
      if (m.a.length !== 1) d.push(`E1 ${tag} : ${m.a.length} destinataire(s) dans « À » au lieu d'un (${m.a.join(", ")})`);
      if (m.cci.length) d.push(`E1 ${tag} : Cci non vide (${m.cci.join(", ")})`);
    }
    if (m.objet !== plan.objet) d.push(`E3 ${tag} : objet « ${m.objet.slice(0, 60)} » au lieu de « ${plan.objet.slice(0, 60)} »`);
    if (m.importance !== (plan.importance ? "haute" : "normale")) d.push(`E4 ${tag} : importance « ${m.importance} » au lieu de « ${plan.importance ? "haute" : "normale"} »`);
    // corps : le message entier, une seule fois, d'un seul tenant ; autour, seulement la signature
    if (r.prevu) {
      const B = T(m.corpsTexte), M = T(r.prevu.texte), S = m.signaturesTexte.map(T).filter(Boolean);
      const i = B.indexOf(M);
      if (!M) d.push(`ORACLE ${tag} : message prévu vide`);
      else if (i < 0) {
        const moitie = M.slice(0, Math.floor(M.length / 2));
        d.push(`E3 ${tag} : message ${B.indexOf(moitie) >= 0 ? "TRONQUÉ ou COUPÉ" : "ABSENT"} dans le corps envoyé (${B.length} lettres envoyées, ${M.length} prévues ; corps : « ${m.corpsTexte.replace(/\s+/g, " ").slice(0, 160)} »)`);
      } else if (B.indexOf(M, i + 1) >= 0) d.push(`E3 ${tag} : message EN DOUBLE dans le corps`);
      else {
        const autour = B.slice(0, i) + "|" + B.slice(i + M.length);
        const permis = plan.sansSignature ? ["|"] : (S.length ? permutations(S).flatMap((x) => [...Array(x.length + 1).keys()].map((k) => x.slice(0, k) + "|" + x.slice(k))) : ["|"]);
        if (!permis.includes(autour)) {
          const sigRestee = plan.sansSignature && S.some((x) => x && autour.replace("|", "").includes(x.slice(0, 8)));
          d.push(`E3 ${tag} : ${sigRestee ? "texte de la signature NON retiré" : "texte EN TROP dans le corps"} (« ${autour.replace("|", " … ").slice(0, 120)} »)`);
        }
      }
      // liens : ceux du message (adresse et texte), plus ceux des signatures, rien d'autre
      const attendus = [...r.prevu.liens.map((l) => sansBarre(l.href)), ...m.signaturesLiens.map(sansBarre)].sort();
      const vus = m.liens.map((l) => sansBarre(l.href)).sort();
      if (JSON.stringify(attendus) !== JSON.stringify(vus)) d.push(`E3 ${tag} : liens envoyés [${vus.join(" | ").slice(0, 300)}] au lieu de [${attendus.join(" | ").slice(0, 300)}]`);
      r.prevu.liens.forEach((l) => {
        if (!m.liens.some((x) => sansBarre(x.href) === sansBarre(l.href) && T(x.texte) === T(l.texte))) d.push(`E3 ${tag} : lien « ${T(l.texte).slice(0, 30)} » → ${l.href} absent ou modifié`);
      });
    }
  });
  servis.forEach((n, a) => { if (n > 1) d.push(`E2 : « ${a} » a reçu ${n} mails`); });

  // suivi : « envoyé » = vraiment parti, et inversement (sinon perte ou double envoi à la reprise)
  Object.entries(r.statuts).forEach(([a, s]) => {
    if (s === "envoye" && !servis.has(a)) d.push(`PERTE : « ${a} » noté envoyé alors que rien n'est parti`);
  });
  servis.forEach((_, a) => {
    const s = r.statuts[a];
    if (prevus.has(a) && s !== "envoye" && s !== "verifier") d.push(`E2 : « ${a} » est parti mais noté « ${s ?? "absent"} » (il serait renvoyé à la reprise)`);
  });
  // refus explicite d'Outlook (alerte « n'a pas pu être envoyé ») : rien n'est parti, le
  // destinataire ne doit pas être classé « à vérifier » (donc écarté pour toujours)
  if (r.journal.some((j) => j.type === "alerte")) {
    Object.entries(r.statuts).forEach(([a, s]) => {
      if (s === "verifier" && !servis.has(a)) d.push(`PERTE : « ${a} » classé « à vérifier » (rangé dans Déjà envoyés) alors qu'Outlook a refusé l'envoi`);
    });
  }
  // liste de l'envoi : chaque adresse prévue, telle quelle, ou un message qui dit pourquoi non
  const liste = new Set(r.liste.map(minus));
  r.liste.forEach((a) => { if (!prevus.has(minus(a))) d.push(`E1 : adresse « ${a} » dans la liste d'envoi alors que l'utilisateur ne l'a jamais écrite`); });
  plan.destinataires.forEach((a) => {
    const dite = r.infos.toLowerCase().includes(minus(a)) || r.notice.toLowerCase().includes(minus(a));
    if (!liste.has(minus(a)) && r.liste.length && !dite) d.push(`SILENCE : adresse prévue « ${a} » absente de la liste d'envoi, sans message`);
  });
  // clics et saisies hors des commandes prévues
  r.journal.filter((j) => j.type === "clic-interdit").forEach((j) => d.push(`E6 : clic sur « ${j.detail.libelle} » (${j.detail.role})`));
  if (r.journal.some((j) => j.type === "saisie-redaction-cachee")) d.push(`E6/R5 : saisie dans une rédaction cachée ou un brouillon (${JSON.stringify(r.journal.filter((j) => j.type === "saisie-redaction-cachee")[0]!.detail)})`);
  // arrêt : sans envoi et avec un message clair
  const restants = plan.destinataires.filter((a) => !servis.has(minus(a)) && r.statuts[minus(a)] !== "deja");
  if (restants.length && r.fini && (!r.notice || r.noticeType !== "warn")) d.push(`MESSAGE : ${restants.length} destinataire(s) non servi(s) sans message d'avertissement (notice « ${r.notice} »)`);
  return [...new Set(d)];
}
