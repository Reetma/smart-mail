// Moteur d'envoi : une campagne (objet, message HTML, liste de destinataires)
// gardée dans le stockage local de l'extension, envoyée un destinataire à la
// fois (ou par paquets en Cci) par les gestes de PP.Outlook, avec une pause au
// hasard entre deux mails.
//
// Stockage pensé pour les grosses listes : la liste (« pp-liste ») est écrite une
// seule fois ; après chaque mail on ne réécrit que la campagne (« pp-campagne ») :
// position (curseur) et résultats des destinataires déjà traités.
//
// Statuts : attente (pas encore traité, ou arrêt avant « Envoyer »), envoi
// (écrit AVANT le clic « Envoyer »), envoye, verifier (issue incertaine : jamais
// renvoyé tout seul), deja (dans la liste « Déjà envoyés » : sauté). Un « envoi »
// retrouvé au chargement devient « verifier ».
//
// Liste « Déjà envoyés » (« pp-deja ») : toutes les adresses servies (envoyées ou
// à vérifier), toutes campagnes confondues. Elles sont retirées des nouveaux
// envois, et recopiées dans le fichier publipostage-deja-envoyes.txt
// (Téléchargements) à chaque arrêt de l'envoi.

(function (root) {
  "use strict";
  var PP = (root.PP = root.PP || {});
  var api = root.chrome || root.browser;
  var O = PP.Outlook;

  var CLE = "pp-campagne";
  var CLE_LISTE = "pp-liste";
  var CLE_ENVOIS = "pp-envois";
  var CLE_ACTIF = "pp-actif";
  var CLE_DEJA = "pp-deja";
  var FICHIER_DEJA = "publipostage-deja-envoyes.txt";
  var JOUR = 24 * 3600 * 1000;

  // cci : un mail pour « paquet » destinataires en Cci (« À » vide) ; maxJour compte
  // alors des destinataires. importance : « Importance haute ». sansTexteSignature :
  // texte de la signature retiré, images gardées.
  var DEFAUTS = { pauseMin: 25, pauseMax: 60, maxJour: 150, cci: false, paquet: 30, importance: false, sansTexteSignature: false };
  var PAQUET_MAX = 500; // limite d'Outlook par message
  function taillePaquet(r) { return Math.max(1, Math.min(PAQUET_MAX, Math.floor(r.paquet) || DEFAUTS.paquet)); }

  function lire(cle) {
    return new Promise(function (r) { try { api.storage.local.get(cle, function (o) { r(o ? o[cle] : undefined); }); } catch (e) { r(undefined); } });
  }
  function ecrire(cle, valeur) {
    return new Promise(function (r, k) {
      try {
        var o = {}; o[cle] = valeur;
        api.storage.local.set(o, function () { var err = api.runtime && api.runtime.lastError; if (err) k(new Error(err.message)); else r(); });
      } catch (e) { k(e); }
    });
  }
  function effacer(cle) {
    return new Promise(function (r) { try { api.storage.local.remove(cle, function () { r(); }); } catch (e) { r(); } });
  }

  // Adresses trouvées dans un texte quelconque : une par ligne (format attendu), ou
  // séparées par ; , espaces, « Nom <adresse> », colonne CSV. Domaine en minuscules,
  // doublons retirés (sans tenir compte de la casse). Tient des fichiers de
  // plusieurs millions de lignes.
  // Une adresse n'est prise que ENTIÈRE, bornée par un séparateur (espace, ; , < > ( ) [ ]
  // guillemets, deux-points…) : « andré.martin@… » ou « a&b@… » ne donnent jamais
  // « martin@… » ou « b@… » (une autre personne). Les candidats rejetés sont comptés.
  var TROUVE = /(?<=^|[\s;,<>()\[\]"':=|])[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}(?=$|[\s;,<>()\[\]"':|.?!])/g;
  var CANDIDAT = /[^\s;,<>()\[\]"']+@[^\s;,<>()\[\]"']+/g;
  function extraireAdresses(texte, dejaVues) {
    var vues = dejaVues || new Set(), valides = [], doublons = 0, m;
    var s = String(texte || "");
    var prises = 0;
    TROUVE.lastIndex = 0;
    while ((m = TROUVE.exec(s))) {
      prises++;
      var a = m[0].replace(/^[.']+|[.']+$/g, "");
      var at = a.lastIndexOf("@");
      a = a.slice(0, at) + "@" + a.slice(at + 1).toLowerCase();
      if (!O.MAIL.test(a)) continue;
      var cle = a.toLowerCase();
      if (vues.has(cle)) { doublons++; continue; }
      vues.add(cle);
      valides.push(a);
    }
    var candidats = (s.match(CANDIDAT) || []).length;
    return { adresses: valides, doublons: doublons, vues: vues, invalides: Math.max(0, candidats - prises) };
  }

  // Pause entière tirée au hasard entre pauseMin et pauseMax, bornes incluses.
  function tirerPause(r) {
    var min = Math.max(0, Math.floor(r.pauseMin)), max = Math.max(min, Math.floor(r.pauseMax));
    return min + Math.floor(Math.random() * (max - min + 1));
  }

  function compter(c) {
    var n = { total: 0, attente: 0, envoye: 0, verifier: 0, deja: 0 };
    if (!c) return n;
    n.total = c.total;
    n.attente = c.total - c.curseur;
    for (var k in c.resultats) {
      var s = c.resultats[k].s;
      if (s === "envoye") n.envoye++;
      else if (s === "verifier") n.verifier++;
      else if (s === "deja") n.deja++;
    }
    return n;
  }

  function Moteur(ui) {
    this.ui = ui;
    this.enCours = false;
    this.arret = false;
    this.campagne = null;
    this.liste = [];
    this.deja = new Set(); // adresses « déjà envoyées », en minuscules
    this.dejaListe = [];
    this.dejaModifie = false;
  }

  // ------------------------------------------------------------ déjà envoyés
  Moteur.prototype.chargerDeja = async function () {
    var l = await lire(CLE_DEJA);
    this.dejaListe = Array.isArray(l) ? l : [];
    this.deja = new Set(this.dejaListe.map(function (a) { return a.toLowerCase(); }));
  };
  Moteur.prototype.estDeja = function (a) { return !!a && this.deja.has(a.toLowerCase()); };
  // Ajoute des adresses (relit d'abord le stockage : un autre onglet a pu en ajouter).
  Moteur.prototype.ajouterDeja = async function (adresses) {
    await this.chargerDeja();
    var n = 0;
    for (var i = 0; i < adresses.length; i++) {
      var a = adresses[i], k = a && a.toLowerCase();
      if (!k || this.deja.has(k)) continue;
      this.deja.add(k);
      this.dejaListe.push(a);
      n++;
    }
    if (n) { await ecrire(CLE_DEJA, this.dejaListe); this.dejaModifie = true; }
    return n;
  };
  Moteur.prototype.viderDeja = async function () {
    this.deja = new Set();
    this.dejaListe = [];
    await effacer(CLE_DEJA);
  };
  Moteur.prototype.texteDeja = function () { return this.dejaListe.length ? this.dejaListe.join("\r\n") + "\r\n" : ""; };
  // Fichier publipostage-deja-envoyes.txt (Téléchargements), remplacé à chaque fois
  // (par le service de l'extension). Trop gros pour cela : téléchargement daté.
  Moteur.prototype.ecrireFichierDeja = async function () {
    var texte = this.texteDeja();
    this.dejaModifie = false;
    if (texte.length < 1500000) {
      try {
        var r = await api.runtime.sendMessage({ pp: "fichier", nom: FICHIER_DEJA, texte: texte });
        if (r && r.ok) return { ok: true, nom: FICHIER_DEJA };
      } catch (e) {}
    }
    var d = new Date(), p = function (x) { return String(x).padStart(2, "0"); };
    var nom = FICHIER_DEJA.replace(".txt", "-" + d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "-" + p(d.getHours()) + p(d.getMinutes()) + ".txt");
    var url = URL.createObjectURL(new Blob([texte], { type: "text/plain;charset=utf-8" }));
    var a = document.createElement("a");
    a.href = url; a.download = nom; a.style.display = "none";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
    return { ok: true, nom: nom };
  };

  Moteur.prototype.charger = async function () {
    await this.chargerDeja();
    var c = await lire(CLE);
    var liste = await lire(CLE_LISTE);
    if (!c || !Array.isArray(liste) || typeof c.curseur !== "number") { this.campagne = null; this.liste = []; return null; }
    // Mail (un destinataire ou un paquet en Cci) interrompu après « envoi » : à vérifier.
    var change = false, convertis = 0;
    for (var r = c.resultats[c.curseur]; r && r.s === "envoi"; r = c.resultats[c.curseur]) {
      c.resultats[c.curseur] = { s: "verifier", t: r.t, d: "Page rechargée ou fermée pendant l'envoi : vérifiez les Éléments envoyés.", c: r.c };
      c.curseur++;
      change = true; convertis++;
    }
    if (change) { await ecrire(CLE, c); await noterEnvoi(convertis); } // issue incertaine : comptée dans la limite du jour
    this.campagne = c;
    this.liste = liste;
    // Adresses déjà servies par cette campagne (versions précédentes comprises).
    var servis = [];
    for (var k in c.resultats) if (c.resultats[k].s === "envoye" || c.resultats[k].s === "verifier") servis.push(liste[+k]);
    await this.ajouterDeja(servis);
    return c;
  };

  Moteur.prototype.creer = async function (objet, html, adresses, reglages) {
    var c = {
      id: Date.now().toString(36),
      creeLe: Date.now(),
      objet: objet,
      html: html,
      reglages: Object.assign({}, DEFAUTS, reglages || {}),
      total: adresses.length,
      curseur: 0,
      resultats: {},
    };
    await ecrire(CLE_LISTE, adresses);
    await ecrire(CLE, c);
    this.campagne = c;
    this.liste = adresses;
    return c;
  };

  Moteur.prototype.abandonner = async function () {
    this.campagne = null;
    this.liste = [];
    await effacer(CLE);
    await effacer(CLE_LISTE);
  };

  Moteur.prototype.reglerRythme = async function (reglages) {
    if (!this.campagne) return;
    if (!this.enCours && await actifAilleurs()) return false; // c'est l'autre onglet qui envoie
    this.campagne.reglages = Object.assign({}, this.campagne.reglages, reglages);
    await ecrire(CLE, this.campagne);
  };
  // Message modifié dans le panneau pendant une campagne : les destinataires encore
  // en attente recevront le nouveau texte (ceux déjà servis ont eu l'ancien).
  Moteur.prototype.majMessage = async function (html) {
    if (!this.campagne) return;
    if (!this.enCours && await actifAilleurs()) return false; // c'est l'autre onglet qui envoie
    this.campagne.html = html;
    await ecrire(CLE, this.campagne);
  };

  // Lignes du suivi : problèmes d'abord, puis envoyés (récents en tête), puis la suite de la file.
  Moteur.prototype.lignes = function (max, suite) {
    var c = this.campagne, liste = this.liste;
    if (!c) return { lignes: [], reste: 0 };
    var ordre = { verifier: 0, envoi: 1, attente: 2, envoye: 3, deja: 4 };
    var faits = Object.keys(c.resultats).map(function (k) { var r = c.resultats[k]; return { i: +k, a: liste[+k], s: r.s, t: r.t, d: r.d, c: r.c }; });
    faits.sort(function (x, y) { return ordre[x.s] - ordre[y.s] || (y.t || 0) - (x.t || 0); });
    var lignes = faits.slice(0, max);
    var vus = {};
    faits.forEach(function (f) { vus[f.i] = 1; });
    var fin = Math.min(c.total, c.curseur + suite);
    for (var i = c.curseur; i < fin; i++) if (!vus[i]) lignes.push({ i: i, a: liste[i], s: "attente" });
    return { lignes: lignes, reste: Math.max(0, c.total - fin) };
  };

  // Toutes les lignes (export) : adresse, statut, date, détail.
  Moteur.prototype.toutes = function () {
    var c = this.campagne, liste = this.liste, out = [];
    if (!c) return out;
    for (var i = 0; i < c.total; i++) {
      var r = c.resultats[i];
      out.push({ a: liste[i], s: r ? r.s : "attente", t: r && r.t, d: r && r.d, c: r && r.c });
    }
    return out;
  };

  // Envois réels des dernières 24 h (toutes campagnes), pour la limite quotidienne.
  Moteur.prototype.envoisRecents = async function () {
    var liste = (await lire(CLE_ENVOIS)) || [];
    var limite = Date.now() - JOUR;
    var garde = liste.filter(function (t) { return t > limite; });
    if (garde.length !== liste.length) await ecrire(CLE_ENVOIS, garde);
    return garde;
  };
  // Un horodatage par destinataire (un mail en Cci à 30 personnes en compte 30).
  async function noterEnvoi(nb) {
    var liste = (await lire(CLE_ENVOIS)) || [];
    for (var i = 0, t = Date.now(); i < (nb || 1); i++) liste.push(t);
    await ecrire(CLE_ENVOIS, liste);
  }

  // Un seul envoi à la fois dans tout le navigateur (plusieurs onglets Outlook) :
  // un marqueur rafraîchi pendant l'envoi, périmé après 30 s sans signe de vie.
  var ID_ONGLET = Math.random().toString(36).slice(2);
  async function actifAilleurs() {
    var m = await lire(CLE_ACTIF);
    return !!(m && m.id !== ID_ONGLET && Date.now() - m.t < 90000);
  }
  function battement() { return ecrire(CLE_ACTIF, { id: ID_ONGLET, t: Date.now() }).catch(function () {}); }
  function liberer() { return effacer(CLE_ACTIF); }
  // Signature d'une campagne indépendante de l'ordre des clés (Chrome les trie à
  // l'enregistrement) : identifiant, curseur, et statut + date de chaque résultat.
  function signature(c) {
    var r = c.resultats || {};
    return [c.id, c.curseur].concat(Object.keys(r).sort(function (x, y) { return x - y; }).map(function (k) { return k + ":" + (r[k] && r[k].s) + ":" + ((r[k] && r[k].t) || ""); })).join("|");
  }
  // L'extension est-elle toujours là (pas rechargée ni désactivée) ?
  function vivante() { try { return !!(api && api.runtime && api.runtime.id); } catch (e) { return false; } }
  Moteur.prototype.envoiAilleurs = function () { return actifAilleurs(); };
  // Envoi en arrière-plan (onglet caché, autre application devant) : pendant
  // l'envoi, Chrome ne décharge pas l'onglet et le Mac ne se met pas en veille
  // (l'écran peut s'éteindre). Voir sw.js.
  function signalerEnvoi(actif) {
    try { return api.runtime.sendMessage({ pp: "envoi", actif: actif }).catch(function () {}); } catch (e) { return Promise.resolve(); }
  }
  try { root.addEventListener("pagehide", function () { if (PP.moteurActif) { liberer(); signalerEnvoi(false); } }); } catch (e) {}

  // Un mail : ouvrir, À (ou un paquet d'adresses en Cci), objet, corps, importance,
  // Envoyer. Renvoie { statut, detail } :
  //   "pret-echec" arrêt avant « Envoyer » (rien n'est parti) ; "envoye" ; "verifier".
  Moteur.prototype.unMail = async function (adresse, objet, html, reglages, avantEnvoi, cci) {
    var ui = this.ui;
    reglages = reglages || {};
    O.surAttente = function (texte) { ui.etape(texte); };
    if (!vivante()) return { statut: "pret-echec", detail: "L'extension a été rechargée ou désactivée : rechargez la page Outlook." };
    if (root.navigator && root.navigator.onLine === false) return { statut: "pret-echec", detail: "Hors connexion : l'envoi est en pause (les mails resteraient dans la Boîte d'envoi). Reprenez quand la connexion revient." };
    ui.etape("Nouveau message");
    var r = await O.ouvrirRedaction();
    if (!r.ok) return { statut: "pret-echec", detail: r.raison };
    if (cci) {
      ui.etape("Cci : " + cci.length + " destinataire(s)");
      r = await O.remplirCci(cci);
    } else {
      ui.etape("Destinataire : " + adresse);
      r = await O.remplirA(adresse);
    }
    if (!r.ok) return { statut: "pret-echec", detail: r.raison };
    ui.etape("Objet");
    r = await O.remplirObjet(objet);
    if (!r.ok) return { statut: "pret-echec", detail: r.raison };
    ui.etape(reglages.sansTexteSignature ? "Message, texte de la signature retiré" : "Message");
    r = await O.remplirCorps(html, { sansTexteSignature: reglages.sansTexteSignature });
    if (!r.ok) return { statut: "pret-echec", detail: r.raison };
    if (reglages.importance) {
      ui.etape("Importance haute");
      r = await O.importanceHaute();
      if (!r.ok) return { statut: "pret-echec", detail: r.raison };
    }
    await O.attendre(O.hasard(600, 1500)); // relecture
    if (reglages.sansTexteSignature) {
      // Une signature posée après coup (complément) est retirée elle aussi.
      r = O.retirerTexteSignature();
      if (!r.ok) return { statut: "pret-echec", detail: r.raison };
    }
    if (root.navigator && root.navigator.onLine === false) return { statut: "pret-echec", detail: "Hors connexion : rien n'est envoyé. Reprenez quand la connexion revient." };
    if (!vivante()) return { statut: "pret-echec", detail: "L'extension a été rechargée ou désactivée : rien n'est envoyé." };
    if (avantEnvoi) await avantEnvoi();
    ui.etape("Envoyer");
    var nb = cci ? cci.length : 1;
    r = await O.envoyer();
    if (r.ok) { await noterEnvoi(nb); return { statut: "envoye" }; }
    if (r.envoye === false) return { statut: "pret-echec", detail: r.raison };
    await noterEnvoi(nb); // issue incertaine : comptée par prudence dans la limite du jour
    return { statut: "verifier", detail: r.raison };
  };

  Moteur.prototype.verifierLimite = async function (maxJour, unite) {
    var recents = await this.envoisRecents();
    if (maxJour > 0 && recents.length >= maxJour) {
      var prochain = new Date(recents[0] + JOUR);
      return "Limite de " + maxJour + " " + (unite || "mails") + " sur 24 h atteinte. Reprise possible après " + prochain.toLocaleString("fr-FR", { weekday: "long", hour: "2-digit", minute: "2-digit" }) + ".";
    }
    return null;
  };

  async function verrouiller(self) {
    if (self.enCours) return "Un envoi est déjà en cours.";
    if (await actifAilleurs()) return "Un envoi est déjà en cours dans un autre onglet Outlook (ou vient de s'arrêter : réessayez dans 2 minutes).";
    self.enCours = true; PP.moteurActif = true; self.arret = false;
    await battement();
    self.minuterie = setInterval(battement, 5000);
    await signalerEnvoi(true);
    return null;
  }
  async function deverrouiller(self) {
    clearInterval(self.minuterie);
    self.enCours = false; PP.moteurActif = false;
    await liberer();
    await signalerEnvoi(false);
  }

  function bilan(c) {
    var n = compter(c);
    var fait = c.mails && c.mails !== n.envoye ? n.envoye + " destinataire(s) servi(s) en " + c.mails + " mail(s)" : n.envoye + " mail(s) envoyé(s)";
    return { message: "Terminé : " + fait + (n.verifier ? ", " + n.verifier + " à vérifier." : ".") + (n.deja ? " " + n.deja + " adresse(s) déjà servie(s) avant : sautée(s)." : ""), type: n.verifier ? "warn" : "ok" };
  }

  // Boucle principale : jusqu'à la fin de la liste, une pause, une limite ou un arrêt.
  Moteur.prototype.lancer = async function () {
    var self = this, c = this.campagne;
    if (!c) return { message: "Aucune campagne.", type: "warn" };
    var occupe = await verrouiller(this);
    if (occupe) return { message: occupe, type: "warn" };
    self.ui.debut();
    try {
      await self.chargerDeja();
      // Pause tenue même après un rechargement : pas de mail avant pauseMin depuis le dernier.
      var recents0 = await self.envoisRecents(), dernier = recents0.length ? recents0[recents0.length - 1] : 0;
      for (var reste0 = Math.ceil((dernier + c.reglages.pauseMin * 1000 - Date.now()) / 1000); reste0 > 0 && !self.arret; reste0--) {
        self.ui.pause(reste0, c.reglages.pauseMin, c.reglages);
        await O.attendre(1000);
        if (reste0 % 10 === 0) battement();
      }
      for (;;) {
        if (self.arret) return { message: "Envoi en pause. « Reprendre l'envoi » continue au destinataire suivant.", type: "" };
        if (!vivante()) return { message: "L'extension a été rechargée ou désactivée : rechargez la page Outlook, puis reprenez. Rien de plus n'est envoyé d'ici.", type: "warn" };
        // L'envoi a-t-il été modifié ailleurs (autre onglet) ? Alors on s'arrête et on recharge.
        var fraiche = await lire(CLE);
        if (!fraiche || signature(fraiche) !== signature(c)) {
          await self.charger();
          return { message: "Cet envoi a été modifié dans un autre onglet : l'état à jour est rechargé. Rien n'a été envoyé d'ici ; relancez si besoin.", type: "warn" };
        }
        await self.chargerDeja(); // « Déjà envoyés » complétés ailleurs pendant la pause
        battement();
        // Adresses de la liste « Déjà envoyés » (importée en cours de route, par exemple) : sautées.
        var sautees = 0;
        while (c.curseur < c.total && self.estDeja(self.liste[c.curseur])) { c.resultats[c.curseur] = { s: "deja", t: Date.now() }; c.curseur++; sautees++; }
        if (sautees) { await ecrire(CLE, c); self.ui.etat(c); }
        if (c.curseur >= c.total) return bilan(c);
        var reglages = c.reglages, cci = !!reglages.cci;
        var limite = await self.verifierLimite(reglages.maxJour, cci ? "destinataires" : "mails");
        if (limite) return { message: limite, type: "warn" };
        // Destinataires de ce mail : un seul, ou jusqu'à « paquet » en Cci (dans la
        // limite du jour), à la suite dans la liste et pas encore servis.
        var taille = 1;
        if (cci) {
          var resteJour = reglages.maxJour > 0 ? reglages.maxJour - (await self.envoisRecents()).length : Infinity;
          taille = Math.max(1, Math.min(taillePaquet(reglages), resteJour));
        }
        var indices = [];
        for (var j = c.curseur; j < c.total && indices.length < taille && !self.estDeja(self.liste[j]); j++) indices.push(j);
        var adresses = indices.map(function (k) { return self.liste[k]; });
        var adresse = adresses[0];
        var nom = cci ? adresses.length + " destinataire(s) en Cci (" + adresse + (adresses.length > 1 ? " … " + adresses[adresses.length - 1] : "") + ")" : adresse;
        self.ui.courant(nom, compter(c));
        var r = await self.unMail(cci ? "" : adresse, c.objet, c.html, reglages, async function () {
          var t = Date.now();
          indices.forEach(function (k) { c.resultats[k] = cci ? { s: "envoi", t: t, c: 1 } : { s: "envoi", t: t }; });
          await ecrire(CLE, c);
        }, cci ? adresses : null);
        var t = Date.now();
        indices.forEach(function (k) {
          var res = r.statut === "envoye" ? { s: "envoye", t: t } : r.statut === "verifier" ? { s: "verifier", t: t, d: r.detail } : { s: "attente", d: r.detail };
          if (cci && r.statut !== "pret-echec") res.c = 1;
          c.resultats[k] = res;
        });
        if (r.statut !== "pret-echec") {
          c.curseur = indices[indices.length - 1] + 1;
          if (r.statut === "envoye") c.mails = (c.mails || 0) + 1;
        }
        await ecrire(CLE, c);
        if (r.statut !== "pret-echec") await self.ajouterDeja(adresses);
        self.ui.etat(c);
        if (r.statut === "verifier") return { message: (cci ? adresses.length + " destinataire(s) en Cci" : adresse) + " : " + r.detail + " Vérifiez les Éléments envoyés, puis reprenez. " + (cci ? "Ces destinataires ne seront pas renvoyés." : "Ce destinataire ne sera pas renvoyé."), type: "warn" };
        if (r.statut === "pret-echec") return { message: "Arrêt avant l'envoi à " + nom + " (rien n'est parti) : " + r.detail, type: "warn" };
        if (c.curseur >= c.total) continue;
        // Pause au hasard entre deux mails (bornes du réglage incluses), décomptée
        // seconde par seconde. Le réglage est relu à chaque seconde : s'il change
        // pendant la pause, celle-ci est retirée dans les nouvelles bornes.
        var borne = { min: c.reglages.pauseMin, max: c.reglages.pauseMax };
        var tiree = tirerPause(c.reglages), ecoulees = 0;
        while (ecoulees < tiree && !self.arret) {
          if (c.reglages.pauseMin !== borne.min || c.reglages.pauseMax !== borne.max) {
            borne = { min: c.reglages.pauseMin, max: c.reglages.pauseMax };
            tiree = tirerPause(c.reglages);
            if (ecoulees >= tiree) break;
          }
          self.ui.pause(tiree - ecoulees, tiree, c.reglages);
          await O.attendre(1000);
          ecoulees++;
          if (ecoulees % 10 === 0) battement();
        }
      }
    } finally {
      await deverrouiller(self);
      if (self.dejaModifie) { try { await self.ecrireFichierDeja(); } catch (e) {} }
    }
  };

  Moteur.DEFAUTS = DEFAUTS;
  Moteur.PAQUET_MAX = PAQUET_MAX;
  Moteur.taillePaquet = taillePaquet;
  Moteur.extraireAdresses = extraireAdresses;
  Moteur.compter = compter;
  PP.Moteur = Moteur;
})(typeof globalThis !== "undefined" ? globalThis : this);
