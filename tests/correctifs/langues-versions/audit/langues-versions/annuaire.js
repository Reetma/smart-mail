// Carnet factice déterministe (pas de hasard) des fausses pages de l'audit
// « Langues et versions ». Même logique de suggestion que tests/fausse-adresses/
// book.js : un contact est proposé si chaque mot tapé commence un mot de son nom
// ou de son adresse ; triés par poids puis ordre du carnet ; au plus CAP lignes.
// Script classique : window.ANNUAIRE (et module.exports pour le test Node).
(function (root) {
  "use strict";
  var CAP = 5;
  var BOOK = [
    ["Martin Dupont", "martin.dupont@ecole-exemple.fr", 90],
    ["Marie Martin", "marie.martin@ecole-exemple.fr", 89],
    ["Antoine Quesnel", "antoine.quesnel@ecole-exemple.fr", 88],
    ["Anne-Laure Gautier", "anne-laure.gautier@ecole-exemple.fr", 87],
    ["Alexandre Delmas", "alexandre.delmas@ecole-exemple.fr", 86],
    ["Amandine Ferrand", "amandine.ferrand@ecole-exemple.fr", 85],
    ["Adele Perrin", "adele.perrin@ecole-exemple.fr", 84],
    ["Baptiste Rousseau", "baptiste.rousseau@ecole-exemple.fr", 83],
    ["Camille Bernard", "camille.bernard@ecole-exemple.fr", 82],
    ["David Lambert", "david.lambert@ecole-exemple.fr", 81],
    ["Emma Girard", "emma.girard@ecole-exemple.fr", 80],
    ["Hugo Fontaine", "hugo.fontaine@ecole-exemple.fr", 79],
    ["Julie Henry", "julie.henry@ecole-exemple.fr", 78],
    ["Laura Simon", "laura.simon@ecole-exemple.fr", 77],
    ["Nicolas Michel", "nicolas.michel@ecole-exemple.fr", 76],
    ["Paul Robert", "paul.robert@ecole-exemple.fr", 75],
    ["Sophie Guerin", "sophie.guerin@ecole-exemple.fr", 74],
    ["Thomas Petit", "thomas.petit@ecole-exemple.fr", 73],
    ["Zoe Fabre", "zoe.fabre@ecole-exemple.fr", 72],
  ];
  // Annuaire des pastilles : une adresse connue s'affiche sous le NOM du contact
  // (comme le vrai Outlook), une inconnue sous l'adresse.
  var NOMS = {};
  BOOK.forEach(function (c) { NOMS[c[1]] = c[0].toUpperCase(); });
  NOMS["alice@exemple.fr"] = "MARTIN Alice";

  function tokens(s) {
    var n = (s == null ? "" : String(s)).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "") + " ";
    var t = [], w = "";
    for (var i = 0; i < n.length; i++) {
      var c = n[i];
      if ((c >= "a" && c <= "z") || (c >= "0" && c <= "9")) w += c;
      else if (w) { t.push(w); w = ""; }
    }
    return t;
  }
  function ask(query) {
    var qw = tokens(query);
    if (!qw.length) return [];
    return BOOK.map(function (c, i) { return { name: c[0], mail: c[1], w: c[2], i: i, toks: tokens(c[0] + " " + c[1]) }; })
      .filter(function (c) { return qw.every(function (word) { return c.toks.some(function (t) { return t.indexOf(word) === 0; }); }); })
      .sort(function (a, b) { return b.w - a.w || a.i - b.i; })
      .slice(0, CAP)
      .map(function (c) { return { name: c.name, mail: c.mail }; });
  }
  function nomDe(adr) { return NOMS[String(adr || "").toLowerCase()] || adr; }
  var out = { CAP: CAP, BOOK: BOOK, ask: ask, tokens: tokens, nomDe: nomDe, adresses: BOOK.map(function (c) { return c[1]; }) };
  if (typeof module !== "undefined" && module.exports) module.exports = out;
  else root.ANNUAIRE = out;
})(typeof globalThis !== "undefined" ? globalThis : this);
