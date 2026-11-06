// Le point d'entrée unique du moteur vers le monde : natif(op, args) => Promise.
// Route chaque geste vers l'adaptateur DOM (AODom) ou « fichiers » (AOIo).
// C'est l'exact équivalent navigateur du pont natif de la version Mac, donc le
// cœur (planif.js / moteur.js) tourne sans la moindre modification.

(function (root) {
  "use strict";

  var D = root.AODom, IO = root.AOIo, api = root.browser || root.chrome;

  async function natif(op, args) {
    args = args || {};
    switch (op) {
      // DOM / présence / rédaction
      case "lire": return D.lire();
      case "taper": return D.taper(args.texte);
      case "effacer": return D.effacer(args.n);
      case "focus": return D.focus(args.verrouiller);
      case "reste": return D.reste();
      case "nettoyer": return D.nettoyer();
      case "vivant": return D.vivant();
      case "enLigne": return D.enLigne();
      case "instantane": return D.instantane();
      case "outlookOuvert": return D.outlookOuvert();
      case "activerOutlook": return D.activerOutlook();
      case "gmailDevant": return D.gmailDevant();
      case "auPremierPlan": return D.auPremierPlan();
      case "nouveauMessage": return D.openCompose();     // Outlook : clic « Nouveau message »
      case "ouvrirUrl": return D.openCompose();           // Gmail : on IGNORE l'URL et on clique « Composer » (pas de navigation)
      case "diagnostic": return D.diagnostic(args.etape);
      case "plusEssai": return D.plusEssai();           // diagnostic : essaie de cliquer « plus de résultats »
      case "diagnosticFichier": {
        var nom = "diagnostic-adresses-" + D.SOURCE + ".txt";
        return IO.enregistrerSous(nom, "", D.diagReport(args.lu));
      }
      // Fichiers (storage) / export / presse-papiers
      case "lireTexte": return IO.lireTexte(args.chemin);
      case "ecrire": return IO.ecrire(args.chemin, args.texte);
      case "ajouter": return IO.ajouter(args.chemin, args.texte);
      case "supprimer": return IO.supprimer(args.chemin);
      case "existe": return IO.existe(args.chemin);
      case "bureau": return IO.bureau();
      case "enregistrerSous": return IO.enregistrerSous(args.nom, args.dossier, args.texte);
      case "copier": return IO.copier(args.texte);
      case "ouvrir": return IO.ouvrir(args.chemin);
      default: throw new Error("op inconnue " + op);
    }
  }

  // On enlève les résultats des anciens passages (clés « adresses-<source>-… »)
  // pour ne pas remplir le quota de 10 Mo, en gardant ceux que `keep` indique
  // (le dernier passage). Limité à CETTE messagerie : on ne touche jamais aux
  // fichiers de l'autre (sinon on supprimerait son dernier passage).
  natif.prune = async function (keep, source) {
    try {
      var mark = "adresses-" + (source || "") + "-";
      var all = await api.storage.local.get(null);
      var del = [];
      Object.keys(all).forEach(function (k) {
        if (k.indexOf(mark) < 0) return;
        if (keep && keep.indexOf(k) >= 0) return;
        del.push(k);
      });
      if (del.length) await api.storage.local.remove(del);
    } catch (e) {}
  };

  root.natif = natif;
})(typeof globalThis !== "undefined" ? globalThis : this);
