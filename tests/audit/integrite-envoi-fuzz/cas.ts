// Cas ciblés (non aléatoires) : une fausse page « nouvel Outlook » de base, réaliste
// et sans piège (structure de la fiche de référence), à laquelle chaque test
// ajoute UNE variante. Le plan d'envoi de base : 2 destinataires, message avec un
// lien, objet simple.
import type { Cfg, Plan } from "./generateur";
import { SIGNATURES } from "./generateur";

export const MESSAGE =
  '<div>Chers étudiants,</div><div><br></div><div>Les inscriptions à la certification IA ferment le <strong>31 octobre</strong>.</div>' +
  '<div><br></div><div>Réservez votre place → <a href="https://reetma.com/ecole">OBTENEZ VOTRE CERTIFICATION</a></div>' +
  "<div><br></div><div>Le programme dure six semaines, à raison de deux heures par semaine.</div><div><br></div><div>Cordialement,<br>Reetma School</div>";

export function cfgBase(o: Partial<Cfg> = {}): Cfg {
  return {
    graine: 4242,
    nouveauSansAria: false,
    enveloppes: 1,
    profondeurRacine: 0,
    profondeurObjet: 2,
    idsChangeants: true,
    affichage: "nom",
    collerDest: "pastilles",
    delai: null,
    remplacement: null,
    echec: null,
    doublon: null,
    ccVisible: false,
    cciVisible: false,
    basculeCci: "texte",
    collerCorps: "html",
    filtreCorps: { styles: true },
    signature: { html: SIGNATURES.simple!, position: "apres" },
    signatureComplement: SIGNATURES.simple!,
    tardifs: [],
    restes: [],
    doublonsCaches: [],
    autreEnvoyer: false,
    rerenduFocus: true,
    fenetreOuverture: 0,
    annuaire: [
      { nom: "DUPONT Jean", adresse: "jean.dupont@exemple.fr" },
      { nom: "MARTIN Alice", adresse: "alice.martin@exemple.fr" },
      { nom: "NGUYEN Zoé", adresse: "zoe.nguyen@exemple.fr" },
    ],
    listeMessages: [{ nom: "Paul MARTIN", adresse: "paul.martin@exemple.fr", objet: "Compte rendu" }],
    ...o,
  };
}

export function planBase(o: Partial<Plan> = {}): Plan {
  const destinataires = o.destinataires ?? ["camille.leroy@exemple.fr", "hugo.blanc@exemple.org"];
  return {
    graine: 4242,
    destinataires,
    saisie: o.saisie ?? destinataires.join("\n"),
    objet: "Certification IA - Avant 31 Oct.",
    html: MESSAGE,
    cci: false,
    paquet: 1,
    importance: false,
    sansSignature: false,
    resume: ["cas ciblé"],
    ...o,
  };
}
