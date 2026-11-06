// Carnet factice PARTAGÉ : la fausse page (navigateur) ET la simulation Node
// (le test) l'utilisent, pour que le menu affiché et la séquence de recherches
// attendue coïncident exactement. tokens() reproduit planif.js (NFD + [a-z0-9]).

export const CAP = 5;

export const BOOK = [
  ["Martin Dupont", "martin.dupont@essca.fr", 90],
  ["Marie Martin", "marie.martin@essca.fr", 89],
  ["Martine Lefevre", "martine.lefevre@essca.eu", 88],
  ["Mathis Moreau", "mathis.moreau@essca.fr", 87],
  ["Maxime Marchand", "maxime.marchand@essca.eu", 86],
  ["Manon Mercier", "manon.mercier@essca.fr", 85],
  ["Antoine Quesnel", "antoine.quesnel@essca.eu", 84],
  ["Anne-Laure Gautier", "anne-laure.gautier@essca.fr", 83],
  ["Alexandre Delmas", "alexandre.delmas@essca.eu", 82],
  ["Amandine Ferrand", "amandine.ferrand@essca.eu", 81],
  ["Adele Perrin", "adele.perrin@essca.fr", 80],
  ["Baptiste Rousseau", "baptiste.rousseau@gmail.com", 79],
  ["Camille Bernard", "camille.bernard@essca.fr", 78],
  ["Christelle Dubois", "christelle.dubois@essca.fr", 77],
  ["David Lambert", "david.lambert@essec.edu", 76],
  ["Emma Girard", "emma.girard@essca.fr", 75],
  ["Fanny Lopez", "fanny.lopez@essca.eu", 74],
  ["Gabriel Roux", "gabriel.roux@essca.fr", 73],
  ["Hugo Fontaine", "hugo.fontaine@essca.eu", 72],
  ["Ines Vincent", "ines.vincent@essca.fr", 71],
  ["Julie Henry", "julie.henry@essca.fr", 70],
  ["Kevin Masson", "kevin.masson@essca.eu", 69],
  ["Laura Simon", "laura.simon@essca.fr", 68],
  ["Nicolas Michel", "nicolas.michel@essca.eu", 66],
  ["Olivia Garcia", "olivia.garcia@essca.fr", 65],
  ["Paul Robert", "paul.robert@essca.eu", 64],
  ["Quentin Faure", "quentin.faure@essca.fr", 63],
  ["Romain Blanc", "romain.blanc@essca.eu", 62],
  ["Sophie Guerin", "sophie.guerin@essca.fr", 61],
  ["Thomas Petit", "thomas.petit@essec.edu", 60],
  ["Ugo Chevalier", "ugo.chevalier@essca.fr", 59],
  ["Valentin Leroy", "valentin.leroy@essca.eu", 58],
  ["Wang Li", "g.wang@essca.fr", 57],
  ["Xavier Noel", "xavier.noel@essca.fr", 56],
  ["Yanis Perez", "yanis.perez@essca.eu", 55],
  ["Zoe Fabre", "zoe.fabre@essca.fr", 54],
  ["Bob Durand", "b25123456@essec.edu", 53],
  ["Alice Durand", "b00654321@essec.edu", 52],
];

export function tokens(s) {
  const n = (s == null ? "" : String(s)).toLowerCase().normalize("NFD").replace(/\p{Mn}/gu, "") + " ";
  const t = []; let w = "";
  for (const c of n) {
    if ((c >= "a" && c <= "z") || (c >= "0" && c <= "9")) w += c;
    else if (w) { t.push(w); w = ""; }
  }
  return t;
}

// Les suggestions pour une recherche : contacts dont un mot commence par chaque
// mot de la recherche, triés par poids décroissant puis ordre du carnet, plafonnés.
export function ask(query) {
  const qw = tokens(query);
  return BOOK
    .map(([name, mail, w], i) => ({ name, mail, w, i, toks: tokens(name + " " + mail) }))
    .filter((c) => qw.every((word) => c.toks.some((t) => t.startsWith(word))))
    .sort((a, b) => b.w - a.w || a.i - b.i)
    .slice(0, CAP)
    .map((c) => ({ name: c.name, mail: c.mail }));
}
