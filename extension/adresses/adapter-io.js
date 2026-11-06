// Adaptateur « fichiers » : le moteur écrit/lit des « chemins » ; sur le web ce
// sont des clés de storage.local pour le journal et les résultats, et les
// exports sont des téléchargements (Blob + lien). Aucun accès disque, rien
// n'est envoyé sur le réseau. Exposé via window.AOIo.
//
// Écriture DIFFÉRÉE (write-behind) : storage.local n'a pas d'« append », et
// relire+réécrire le fichier entier à chaque recherche serait en O(n²) (mortel à
// des dizaines de milliers d'adresses). On garde un cache mémoire par chemin
// (base + morceaux ajoutés), on ajoute en O(1), et on écrit en storage de façon
// GROUPÉE : par seuil de volume (256 Ko) et par minuterie de sécurité (2 s), plus
// un flush complet sur ecrire/supprimer (fin de passage) et sur pagehide.
// lireTexte lit TOUJOURS le cache quand il existe, donc reprise et relecture de
// fin voient les ajouts non encore écrits.

(function (root) {
  "use strict";

  var api = root.browser || root.chrome;

  var FLUSH_BYTES = 256 * 1024; // volume ajouté avant d'écrire
  var FLUSH_MS = 2000;          // minuterie de sécurité

  var cache = Object.create(null);     // chemin -> { base, chunks:[] }
  var pending = Object.create(null);   // chemin -> octets depuis le dernier flush
  var timers = Object.create(null);    // chemin -> minuterie
  var inFlight = Object.create(null);  // chemin -> flush en cours
  var reflush = Object.create(null);   // chemin -> re-flush demandé pendant un flush

  function rawRead(p) {
    return api.storage.local.get(p).then(function (o) { return o[p] == null ? null : o[p]; }, function () { return null; });
  }
  function rawWrite(p, v) { var o = {}; o[p] = v; return api.storage.local.set(o); }

  // Écrit la valeur complète courante. Sérialisé par chemin : si un flush est déjà
  // en cours, on demande un re-flush (coalescé) au lieu d'écrire en parallèle.
  function flush(p) {
    if (timers[p]) { clearTimeout(timers[p]); timers[p] = null; }
    if (!cache[p]) return Promise.resolve();
    if (inFlight[p]) { reflush[p] = true; return inFlight[p]; }
    inFlight[p] = (async function () {
      do {
        reflush[p] = false;
        var c = cache[p]; if (!c) break;
        var full = c.chunks.length ? c.base + c.chunks.join("") : c.base;
        c.base = full; c.chunks = []; pending[p] = 0;
        try { await rawWrite(p, full); } catch (e) {}
      } while (reflush[p]);
    })();
    var done = inFlight[p];
    done.then(function () { inFlight[p] = null; }, function () { inFlight[p] = null; });
    return done;
  }
  function flushAll() {
    var keys = Object.keys(cache), jobs = [];
    for (var i = 0; i < keys.length; i++) jobs.push(flush(keys[i]));
    return Promise.all(jobs);
  }

  async function lireTexte(chemin) {
    var c = cache[chemin];
    if (c) return c.chunks.length ? c.base + c.chunks.join("") : c.base;
    return rawRead(chemin);
  }
  async function ecrire(chemin, texte) {
    cache[chemin] = { base: String(texte == null ? "" : texte), chunks: [] };
    pending[chemin] = 0;
    if (timers[chemin]) { clearTimeout(timers[chemin]); timers[chemin] = null; }
    // Flush de TOUT : la fin de passage réécrit le journal → ça force aussi
    // l'écriture des .txt/.csv encore en tampon, avant que la fin ne les relise.
    await flushAll();
    return true;
  }
  async function ajouter(chemin, texte) {
    var t = String(texte == null ? "" : texte);
    if (!cache[chemin]) { var cur = await rawRead(chemin); cache[chemin] = { base: cur || "", chunks: [] }; }
    cache[chemin].chunks.push(t);
    pending[chemin] = (pending[chemin] || 0) + t.length;
    if (pending[chemin] >= FLUSH_BYTES) { flush(chemin); }          // sans await : hors du chemin critique
    else {
      // Sans réarmer : le journal est écrit au plus FLUSH_MS après la première
      // ligne en attente, même si les recherches s'enchaînent (plantage = peu perdu).
      if (!timers[chemin]) timers[chemin] = setTimeout(function () { timers[chemin] = null; flush(chemin); }, FLUSH_MS);
      try { if (timers[chemin] && timers[chemin].unref) timers[chemin].unref(); } catch (e) {} // ne bloque pas Node en test
    }
    return true;
  }
  async function supprimer(chemin) {
    if (timers[chemin]) { clearTimeout(timers[chemin]); timers[chemin] = null; }
    // On retire d'abord ce chemin du cache pour que le flushAll suivant ne le
    // réécrive pas, PUIS on flushe le reste : la fin de passage « terminé »
    // supprime le journal, et c'est le seul moment où .txt/.csv sont forcés.
    delete cache[chemin]; pending[chemin] = 0;
    await flushAll();
    if (inFlight[chemin]) { try { await inFlight[chemin]; } catch (e) {} } // pas de suppression sous une écriture en vol
    try { await api.storage.local.remove(chemin); } catch (e) {}
    return true;
  }

  try { root.addEventListener("pagehide", function () { try { flushAll(); } catch (e) {} }); } catch (e) {}

  function telecharger(nom, texte) {
    var blob = new Blob([texte], { type: "text/plain;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = nom; a.style.display = "none";
    document.body.appendChild(a); a.click();
    setTimeout(function () { try { document.body.removeChild(a); } catch (e) {} URL.revokeObjectURL(url); }, 10000);
    return nom;
  }

  async function copier(texte) {
    try { await navigator.clipboard.writeText(texte); return true; }
    catch (e) {
      try {
        var ta = document.createElement("textarea");
        ta.value = texte; ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta); ta.focus(); ta.select();
        var ok = document.execCommand("copy"); document.body.removeChild(ta); return ok;
      } catch (e2) { return false; }
    }
  }

  root.AOIo = {
    lireTexte: lireTexte,
    ecrire: ecrire,
    ajouter: ajouter,
    supprimer: supprimer,
    flushAll: flushAll, // exposé pour les tests / la fermeture
    existe: function () { return true; },
    bureau: function () { return ""; },
    enregistrerSous: function (nom, dossier, texte) { return telecharger(nom, texte); },
    copier: copier,
    ouvrir: async function (chemin) { var t = await lireTexte(chemin); if (t != null) telecharger(String(chemin).split("/").pop(), t); return true; }
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
