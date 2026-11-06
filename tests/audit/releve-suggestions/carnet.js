// Carnet factice PARTAGÉ par les fausses pages (navigateur) et les tests (Node) :
// la liste affichée et ce que le test attend viennent de la même fonction.
// tokens() reproduit planif.js (minuscules, accents retirés, mots [a-z0-9]).

export function tokens(s) {
  const n = (s == null ? "" : String(s)).toLowerCase().normalize("NFD").replace(/\p{Mn}/gu, "") + " ";
  const t = []; let w = "";
  for (const c of n) {
    if ((c >= "a" && c <= "z") || (c >= "0" && c <= "9")) w += c;
    else if (w) { t.push(w); w = ""; }
  }
  return t;
}

// Petit carnet (relevés courts, adresses toutes différentes et faciles à vérifier).
// type : personne | groupe | liste | salle. affichage : nom tel qu'Outlook l'écrit.
export const PETIT = [
  { nom: "Martin Dupont", mail: "martin.dupont@ecole-test.fr", poids: 90 },
  { nom: "Marie Lefort", mail: "marie.lefort@ecole-test.fr", poids: 89 },
  { nom: "Mathis Moreau", mail: "mathis.moreau@ecole-test.fr", poids: 88 },
  { nom: "Manon Mercier", mail: "manon.mercier@ecole-test.fr", poids: 87 },
  { nom: "Maxime Marchand", mail: "maxime.marchand@ecole-test.fr", poids: 86 },
  { nom: "Mila Morel", mail: "mila.morel@ecole-test.fr", poids: 85 },
  { nom: "Antoine Quesnel", mail: "antoine.quesnel@ecole-test.fr", poids: 84 },
  { nom: "Baptiste Rousseau", mail: "baptiste.rousseau@partenaire-test.com", poids: 79, externe: true },
  { nom: "Camille Bernard", mail: "camille.bernard@ecole-test.fr", poids: 78 },
  { nom: "Emma Girard", mail: "emma.girard@ecole-test.fr", poids: 75 },
  { nom: "Hugo Fontaine", mail: "hugo.fontaine@ecole-test.fr", poids: 72 },
  { nom: "Zoe Fabre", mail: "zoe.fabre@ecole-test.fr", poids: 54 },
];

// Le même carnet + des entrées qui ne sont PAS des personnes (vu dans les vrais
// Outlook d'établissement : groupes Microsoft 365, listes de diffusion, salles).
export const AVEC_GROUPES = PETIT.concat([
  { nom: "Tous les étudiants M1", mail: "tous-m1@ecole-test.fr", poids: 83, type: "liste" },
  { nom: "Équipe Marketing", mail: "equipe-marketing@ecole-test.fr", poids: 82, type: "groupe" },
  { nom: "Salle A101 (30)", mail: "salle-a101@ecole-test.fr", poids: 81, type: "salle" },
]);

// Les suggestions pour une recherche : contacts dont un mot (nom + adresse)
// commence par chaque mot de la recherche, par poids décroissant puis ordre du
// carnet. Pas de plafond ici : la page affiche les `cap` premiers.
export function ask(book, query) {
  const qw = tokens(query);
  if (!qw.length) return [];
  return book
    .map((c, i) => ({ ...c, i, toks: tokens((c.affichage || c.nom) + " " + c.mail) }))
    .filter((c) => qw.every((word) => c.toks.some((t) => t.startsWith(word))))
    .sort((a, b) => (b.poids || 0) - (a.poids || 0) || a.i - b.i)
    .map(({ toks, i, ...c }) => c);
}

// Toutes les adresses d'un carnet (minuscules, triées).
export function adresses(book, filtre) {
  return book.filter(filtre || (() => true)).map((c) => c.mail.toLowerCase()).sort();
}
