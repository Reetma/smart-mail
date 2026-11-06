// Smart Mail. Connexion → objectif → analyse de l'annuaire (rien n'est montré) →
// les étudiants arrivent jusqu'à l'objectif → message et paramètres → envoi.
// Tout est enregistré dans le compte : on reprend sur n'importe quel appareil.
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };
  var CLE = "smartmail";
  var PRESETS = {
    prudent: { taille_pack: 1, packs_heure: 15, max_jour: 150, heure_debut: 8, heure_fin: 20 },
    standard: { taille_pack: 1, packs_heure: 35, max_jour: 400, heure_debut: 8, heure_fin: 20 },
    soutenu: { taille_pack: 25, packs_heure: 6, max_jour: 1000, heure_debut: 8, heure_fin: 20 }
  };
  var etat = {
    config: { demo: false }, code: "", compte: "", nom: "",
    objectif: 500, logique: null, resume: "", trouves: [], vus: new Set(), aEnregistrer: [],
    ecartes: 0, pause: false, enCours: false, fini: false, depart: 0, pageDepart: 0,
    tampon: [], // résultats triés pas encore affichés (repris après une pause)
    parametres: Object.assign({ preset: "standard" }, PRESETS.standard), retourParametres: "objectif"
  };

  function sauverLocal() { try { localStorage.setItem(CLE, JSON.stringify({ code: etat.code })); } catch (e) {} }
  function chargerLocal() { try { etat.code = (JSON.parse(localStorage.getItem(CLE) || "{}").code) || ""; } catch (e) {} }
  function nb(n) { return Number(n).toLocaleString("fr-FR"); }
  function dormir(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  var ETAPE = { connexion: 0, objectif: 1, lancement: 1, recolte: 1, envoi: 2, envoye: 3 };
  function ecran(nom) {
    document.querySelectorAll(".ecran").forEach(function (e) { e.hidden = e.id !== "e-" + nom; });
    if (nom in ETAPE) {
      var n = ETAPE[nom];
      document.body.dataset.etape = String(n);
      document.querySelectorAll(".etapes-nav li").forEach(function (li) {
        var k = Number(li.dataset.n);
        li.classList.toggle("active", k === n);
        li.classList.toggle("faite", k < n);
      });
      etat.ecranCourant = nom;
    }
    var connecte = nom !== "connexion";
    $("compte-barre").textContent = connecte ? etat.compte : "";
    $("ouvrir-parametres").hidden = !connecte || nom === "parametres";
    window.scrollTo(0, 0);
  }

  async function api(chemin, corps, methode) {
    var r = await fetch(chemin, {
      method: methode || (corps ? "POST" : "GET"),
      headers: Object.assign({ "Content-Type": "application/json" }, etat.code ? { Authorization: "Bearer " + etat.code } : {}),
      body: corps ? JSON.stringify(corps) : undefined
    });
    if (!r.ok) {
      var detail = ""; try { var j = await r.json(); detail = typeof j.detail === "string" ? j.detail : ""; } catch (e) {}
      var err = new Error(detail || ("Erreur " + r.status)); err.statut = r.status; throw err;
    }
    return r.json();
  }

  // ------------------------------------------------------------ connexion
  $("form-demo").onsubmit = async function (ev) {
    ev.preventDefault();
    var bouton = this.querySelector("button"); bouton.disabled = true;
    try {
      var s = await api("/demo/session", { email: $("email-demo").value.trim() });
      etat.code = s.code; sauverLocal();
      await ouvrirCompte();
    } catch (e) { alert("Connexion impossible : " + e.message); }
    bouton.disabled = false;
  };
  $("connexion-microsoft").onclick = function () { alert("La connexion Microsoft sera disponible prochainement."); };

  // Le compte : identité, destinataires déjà trouvés, recherche en cours, paramètres.
  async function ouvrirCompte() {
    var c = await api("/v1/compte");
    etat.compte = c.email || c.nom; etat.nom = c.nom;
    etat.parametres = Object.assign({ preset: "standard" }, PRESETS.standard, c.parametres || {});
    etat.recherche = c.recherche; etat.nbEnregistres = c.destinataires;
    versObjectif();
  }

  // ------------------------------------------------------------ objectif / reprise
  function versObjectif() {
    $("bonjour").textContent = "Connecté en tant que " + etat.compte;
    var r = etat.recherche, n = etat.nbEnregistres || 0;
    $("reprise").hidden = !(r || n);
    if (r || n) {
      var obj = r ? r.objectif : n;
      $("reprise-texte").textContent = nb(n) + " étudiants identifiés" + (r ? " sur un objectif de " + nb(obj) : "") +
        (n >= obj ? ". Objectif atteint." : ".");
      $("reprendre").textContent = n >= obj ? "Voir les destinataires" : "Reprendre";
      $("objectif").value = obj;
    } else {
      $("objectif").value = etat.objectif;
    }
    majPuces();
    ecran("objectif");
  }
  function lireObjectif() {
    var v = parseInt($("objectif").value, 10);
    if (!isFinite(v)) v = 500;
    return Math.max(10, Math.min(20000, v));
  }
  function majPuces() {
    var v = lireObjectif();
    document.querySelectorAll("#raccourcis .puce").forEach(function (p) { p.classList.toggle("actif", Number(p.dataset.v) === v); });
  }
  $("objectif").addEventListener("input", majPuces);
  $("moins").onclick = function () { $("objectif").value = Math.max(10, lireObjectif() - (lireObjectif() > 1000 ? 500 : 50)); majPuces(); };
  $("plus").onclick = function () { $("objectif").value = Math.min(20000, lireObjectif() + (lireObjectif() >= 1000 ? 500 : 50)); majPuces(); };
  document.querySelectorAll("#raccourcis .puce").forEach(function (p) { p.onclick = function () { $("objectif").value = p.dataset.v; majPuces(); }; });

  function erreur(e) {
    alert("Une erreur est survenue : " + e.message);
    if (e.statut === 401) { etat.code = ""; sauverLocal(); ecran("connexion"); } else versObjectif();
  }

  $("start").onclick = function () {
    if (etat.nbEnregistres && !confirm("Une nouvelle recherche remplace les " + nb(etat.nbEnregistres) + " destinataires déjà identifiés. Continuer ?")) return;
    nouvelleRecherche().catch(erreur);
  };
  $("nouvelle").onclick = function () { $("reprise").hidden = true; etat.recherche = null; $("objectif").focus(); };
  $("reprendre").onclick = function () { reprendre().catch(erreur); };

  function remiseAZero() {
    etat.trouves = []; etat.vus = new Set(); etat.aEnregistrer = []; etat.ecartes = 0; etat.depart = 0; etat.pageDepart = 0;
    etat.logique = null; etat.resume = ""; etat.pause = false; etat.fini = false; etat.tampon = [];
    $("liste").textContent = "";
  }

  async function nouvelleRecherche() {
    etat.objectif = lireObjectif();
    remiseAZero();
    await api("/v1/destinataires", null, "DELETE");
    etat.nbEnregistres = 0; etat.recherche = null;
    await lancer();
  }

  // Reprise (même sur un autre appareil) : liste et logique relues depuis le compte.
  async function reprendre() {
    remiseAZero();
    var r = etat.recherche || {};
    etat.objectif = r.objectif || lireObjectif();
    etat.depart = r.depart == null ? 0 : r.depart; etat.fini = !!r.fini || r.depart == null; etat.ecartes = r.ecartes || 0;
    etat.logique = r.logique || null; etat.resume = r.resume || "";
    for (var d = 0; ; d += 2000) {
      var l = await api("/v1/destinataires?depart=" + d + "&taille=2000");
      l.personnes.forEach(function (p) { etat.trouves.push(p); etat.vus.add(p.adresse); });
      if (l.personnes.length < 2000) break;
    }
    etat.trouves.slice(-300).forEach(function (p) { ajouter(p, p.source, true); });
    if (!etat.logique && etat.trouves.length < etat.objectif) { await lancer(); return; }
    entrerRecolte();
    await recolter();
  }

  // ------------------------------------------------------------ source des contacts
  // Démonstration : faux annuaire du serveur. Version réelle : annuaire Microsoft (Graph).
  var ECOLE = new URLSearchParams(location.search).get("ecole") || "prenom-nom"; // démo : école fictive
  function pageAnnuaire(depart, taille) {
    return api("/demo/annuaire?depart=" + depart + "&taille=" + taille + "&ecole=" + encodeURIComponent(ECOLE));
  }

  // Échantillon d'environ 30 personnes, en privilégiant celles qui ont un intitulé
  // (poste, programme) : ce sont elles qui révèlent la logique de l'école.
  function echantillon(personnes) {
    var avec = personnes.filter(function (p) { return p.intitule; }), sans = personnes.filter(function (p) { return !p.intitule; });
    return avec.slice(0, 22).concat(sans.slice(0, 30 - Math.min(22, avec.length))).slice(0, 30);
  }

  // ------------------------------------------------------------ analyse
  function etape(n, classe) { document.querySelector('#etapes li[data-e="' + n + '"]').className = classe || ""; }

  async function lancer() {
    ecran("lancement");
    [1, 2, 3].forEach(function (n) { etape(n, ""); });
    var t0 = Date.now();
    etape(1, "encours");
    var premiere = await pageAnnuaire(0, 200);
    await dormir(Math.max(0, 900 - (Date.now() - t0)));
    etape(1, "fait"); etape(2, "encours");
    var t1 = Date.now();
    var appris = await api("/v1/logique", { echantillon: echantillon(premiere.personnes) });
    etat.logique = appris.logique; etat.resume = appris.resume || "";
    await dormir(Math.max(0, 1100 - (Date.now() - t1)));
    etape(2, "fait"); etape(3, "encours");
    await dormir(800);
    etape(3, "fait");
    await enregistrer();
    await dormir(350);
    entrerRecolte();
    await recolter();
  }

  function entrerRecolte() {
    $("logique-resume").textContent = etat.resume;
    $("objectif-rappel").textContent = nb(etat.objectif);
    majCompteurs();
    $("vers-envoi").hidden = true; $("pause").hidden = false; $("pause").textContent = "Suspendre";
    etat.pause = false;
    statut("Recherche en cours", true);
    ecran("recolte");
  }

  // ------------------------------------------------------------ recherche
  function statut(texte, actif) { $("statut-texte").textContent = texte; $("pouls").classList.toggle("arret", !actif); }
  function majCompteurs() { $("trouves").textContent = nb(etat.trouves.length); $("ecartes").textContent = nb(etat.ecartes); }
  var METHODE = { regle: "Règle", logique: "Format", ia: "IA", cache: "Mémoire" };
  function ajouter(p, source, calme) {
    var li = document.createElement("li"); if (!calme) li.className = "nouveau";
    var n = document.createElement("span"); n.className = "n"; n.textContent = p.nom || p.adresse.split("@")[0];
    var a = document.createElement("span"); a.className = "a"; a.textContent = p.adresse;
    var m = document.createElement("span"); m.className = "source"; m.textContent = METHODE[source] || "IA";
    li.appendChild(n); li.appendChild(a); li.appendChild(m);
    var l = $("liste"); l.insertBefore(li, l.firstChild);
    while (l.children.length > 300) l.removeChild(l.lastChild);
  }

  // Enregistre dans le compte les nouveaux destinataires et l'endroit où reprendre.
  async function enregistrer() {
    var lot = etat.aEnregistrer; etat.aEnregistrer = [];
    for (var k = 0; k < lot.length; k += 2000) await api("/v1/destinataires", { personnes: lot.slice(k, k + 2000) });
    etat.nbEnregistres = etat.trouves.length;
    etat.recherche = { objectif: etat.objectif, depart: etat.tampon.length ? etat.pageDepart : etat.depart, fini: etat.fini && !etat.tampon.length,
                       ecartes: etat.ecartes, logique: etat.logique, resume: etat.resume };
    await api("/v1/recherche", etat.recherche, "PUT");
  }

  async function recolter() {
    if (etat.enCours) return;
    etat.enCours = true;
    try {
      while (!etat.pause && etat.trouves.length < etat.objectif && (etat.tampon.length || !etat.fini)) {
        if (!etat.tampon.length) {
          etat.pageDepart = etat.depart;
          var page = await pageAnnuaire(etat.depart, 200);
          var parAdresse = {};
          page.personnes.forEach(function (p) { parAdresse[p.adresse.toLowerCase()] = p; });
          var r = await api("/v1/tri", { contacts: page.personnes, logique: etat.logique });
          etat.tampon = r.resultats.map(function (x) { return { x: x, p: parAdresse[x.adresse] || { adresse: x.adresse, nom: "" } }; });
          etat.depart = page.suivant == null ? null : page.suivant;
          if (page.suivant == null) etat.fini = true;
        }
        // Une adresse à la fois ; une suspension arrête net, le reste du paquet attend.
        var k = 0;
        while (etat.tampon.length && !etat.pause && etat.trouves.length < etat.objectif) {
          var e = etat.tampon.shift();
          if (e.x.groupe === "etudiant") {
            if (etat.vus.has(e.x.adresse)) continue; // déjà dans le compte (reprise)
            var p = { adresse: e.x.adresse, nom: e.p.nom || "", source: e.x.source };
            etat.vus.add(p.adresse); etat.trouves.push(p); etat.aEnregistrer.push(p);
            ajouter(p, p.source);
            majCompteurs();
            await dormir(etat.objectif > 1500 ? 4 : 18); // les adresses arrivent une à une
          } else {
            etat.ecartes++;
            if (++k % 10 === 0) majCompteurs();
          }
        }
        majCompteurs();
        await enregistrer();
      }
    } finally { etat.enCours = false; }
    if (etat.trouves.length >= etat.objectif) {
      statut("Objectif atteint", false); $("pause").hidden = true; $("vers-envoi").hidden = false;
    } else if (etat.fini && !etat.tampon.length) {
      statut("Annuaire entièrement parcouru — " + nb(etat.trouves.length) + " étudiants identifiés", false);
      $("pause").hidden = true; $("vers-envoi").hidden = !etat.trouves.length;
    }
  }

  $("pause").onclick = function () {
    etat.pause = !etat.pause;
    this.textContent = etat.pause ? "Reprendre" : "Suspendre";
    statut(etat.pause ? "Suspendue" : "Recherche en cours", !etat.pause);
    $("vers-envoi").hidden = !(etat.pause && etat.trouves.length);
    if (etat.pause) enregistrer().catch(function () {}); else recolter().catch(erreur);
  };

  // ------------------------------------------------------------ paramètres d'envoi
  function capaciteJour(p) { return Math.min(p.max_jour, p.taille_pack * p.packs_heure * Math.max(1, p.heure_fin - p.heure_debut)); }
  function texteMode(p) { return p.taille_pack === 1 ? "Un message par destinataire" : "Messages groupés de " + p.taille_pack + " (copie cachée)"; }
  function texteRythme(p) {
    return p.packs_heure + " envoi" + (p.packs_heure > 1 ? "s" : "") + " par heure, de " + p.heure_debut + " h à " + p.heure_fin + " h, " +
      nb(p.max_jour) + " destinataires par jour au plus";
  }
  function alertes(p) {
    var a = [];
    if (p.taille_pack > 50) a.push("Au-delà de 50 destinataires par message, les filtres anti-spam bloquent souvent l'envoi.");
    if (p.max_jour > 1000) a.push("Au-delà de 1 000 destinataires par jour, votre compte risque d'être bloqué par le service informatique.");
    if (p.taille_pack * p.packs_heure > 500) a.push("Plus de 500 destinataires par heure : les envois risquent d'être retardés par votre messagerie.");
    return a;
  }
  function presetDe(p) {
    for (var k in PRESETS) {
      var q = PRESETS[k];
      if (q.taille_pack === p.taille_pack && q.packs_heure === p.packs_heure && q.max_jour === p.max_jour && q.heure_debut === p.heure_debut && q.heure_fin === p.heure_fin) return k;
    }
    return "perso";
  }
  function remplirParametres(p) {
    $("p-taille").value = p.taille_pack; $("p-heure").value = p.packs_heure; $("p-jour").value = p.max_jour;
    $("p-debut").value = p.heure_debut; $("p-fin").value = p.heure_fin;
    majParametres();
  }
  function lireParametres() {
    var v = function (id, min, max, def) { var x = parseInt($(id).value, 10); return isFinite(x) ? Math.max(min, Math.min(max, x)) : def; };
    var p = { taille_pack: v("p-taille", 1, 100, 1), packs_heure: v("p-heure", 1, 120, 30), max_jour: v("p-jour", 1, 10000, 400),
              heure_debut: v("p-debut", 0, 23, 8), heure_fin: v("p-fin", 1, 24, 20) };
    if (p.heure_fin <= p.heure_debut) p.heure_fin = Math.min(24, p.heure_debut + 1);
    p.preset = presetDe(p);
    return p;
  }
  function majParametres() {
    var p = lireParametres();
    document.querySelectorAll("#presets .option").forEach(function (o) { o.classList.toggle("actif", o.dataset.preset === p.preset); });
    $("capacite").textContent = "Capacité : " + nb(capaciteJour(p)) + " destinataires par jour. " + texteMode(p) + ".";
    var a = alertes(p);
    $("alerte-param").hidden = !a.length; $("alerte-param").textContent = a.join(" ");
  }
  ["p-taille", "p-heure", "p-jour", "p-debut", "p-fin"].forEach(function (id) { $(id).addEventListener("input", majParametres); });
  document.querySelectorAll("#presets .option").forEach(function (o) {
    o.onclick = function () { remplirParametres(Object.assign({}, PRESETS[o.dataset.preset])); };
  });
  function ouvrirParametres(retour) {
    etat.retourParametres = retour || etat.ecranCourant || "objectif";
    remplirParametres(etat.parametres);
    $("p-email").textContent = etat.compte;
    ecran("parametres");
  }
  $("ouvrir-parametres").onclick = function () { ouvrirParametres(); };
  $("modifier-param").onclick = function () { ouvrirParametres("envoi"); };
  $("retour-parametres").onclick = function () { revenir(); };
  function revenir() { if (etat.retourParametres === "envoi") majRecap(); ecran(etat.retourParametres); }
  async function sauverParametres(p) {
    etat.parametres = await api("/v1/parametres", p, "PUT");
  }
  $("enregistrer-parametres").onclick = async function () {
    try { await sauverParametres(lireParametres()); revenir(); } catch (e) { alert("Enregistrement impossible : " + e.message); }
  };
  $("deconnexion").onclick = async function () {
    try { await api("/v1/deconnexion", {}); } catch (e) {}
    etat.code = ""; sauverLocal(); remiseAZero(); ecran("connexion");
  };

  // ------------------------------------------------------------ message
  $("vers-envoi").onclick = function () {
    etat.pause = true;
    enregistrer().catch(function () {});
    $("nb-dest").textContent = nb(etat.trouves.length);
    majRecap();
    ecran("envoi");
  };
  document.querySelectorAll("#rythme .option").forEach(function (o) {
    o.onclick = function () {
      var p = Object.assign({ preset: o.dataset.preset }, PRESETS[o.dataset.preset]);
      etat.parametres = p; majRecap();
      sauverParametres(p).catch(function () {});
    };
  });
  function majRecap() {
    var p = etat.parametres, n = etat.trouves.length, cap = capaciteJour(p), jours = Math.ceil(n / Math.max(1, cap));
    document.querySelectorAll("#rythme .option").forEach(function (o) { o.classList.toggle("actif", o.dataset.preset === p.preset); });
    $("resume-param").textContent = (p.preset === "perso" ? "Personnalisé : " : "") + texteRythme(p) + ".";
    $("r-dest").textContent = nb(n) + " étudiants";
    $("r-mode").textContent = texteMode(p);
    $("r-rythme").textContent = nb(cap) + " destinataires par jour";
    $("r-duree").textContent = jours <= 1 ? "Moins d'une journée" : jours + " jours";
    $("r-exp").textContent = etat.compte;
  }
  $("retour-recolte").onclick = function () { ecran("recolte"); };
  $("lancer-envoi").onclick = function () {
    if (!$("objet").value.trim() || !$("message").value.trim()) { alert("Indiquez un objet et un message."); return; }
    var n = etat.trouves.length, cap = capaciteJour(etat.parametres);
    $("envoyes").textContent = "0"; $("restants").textContent = nb(n);
    var jours = Math.ceil(n / Math.max(1, cap));
    var fin = new Date(Date.now() + jours * 86400000);
    $("fin-estimee").textContent = jours <= 1 ? "Aujourd'hui" : fin.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });
    document.querySelectorAll(".demo-seulement").forEach(function (e) { e.hidden = !etat.config.demo; });
    ecran("envoye");
    // Démonstration : le compteur avance, rien n'est envoyé.
    var envoyes = 0, pas = etat.parametres.taille_pack;
    clearInterval(etat.minuterie);
    etat.minuterie = setInterval(function () {
      if (envoyes >= n) { clearInterval(etat.minuterie); return; }
      envoyes = Math.min(n, envoyes + pas); $("envoyes").textContent = nb(envoyes); $("restants").textContent = nb(n - envoyes);
    }, 1500);
  };
  $("recommencer").onclick = function () { clearInterval(etat.minuterie); etat.recherche = null; etat.nbEnregistres = etat.trouves.length; versObjectif(); };

  // ------------------------------------------------------------ installation
  // Chrome, Edge, Android : bouton d'installation natif. iPhone/iPad et Safari Mac :
  // pas de bouton automatique, on montre les gestes à faire.
  var invite = null;
  var autonome = window.matchMedia("(display-mode: standalone)").matches || window.navigator.standalone === true;
  var ios = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  var safariMac = !ios && /^((?!chrome|android|crios|fxios|edg).)*safari/i.test(navigator.userAgent);
  window.addEventListener("beforeinstallprompt", function (e) { e.preventDefault(); invite = e; if (!autonome) $("installer").hidden = false; });
  window.addEventListener("appinstalled", function () { $("installer").hidden = true; });
  if (!autonome && (ios || safariMac)) $("installer").hidden = false;
  if (safariMac) {
    var ol = document.querySelector("#feuille-ios ol");
    ol.textContent = "";
    ["Dans la barre des menus, ouvrez Fichier.", "Choisissez « Ajouter au Dock ».", "Smart Mail s'ouvre ensuite comme une application."].forEach(function (t) {
      var li = document.createElement("li"); li.textContent = t; ol.appendChild(li);
    });
  }
  $("installer").onclick = async function () {
    if (invite) { invite.prompt(); try { await invite.userChoice; } catch (e) {} invite = null; $("installer").hidden = true; return; }
    $("feuille-ios").hidden = false;
  };
  $("fermer-feuille").onclick = function () { $("feuille-ios").hidden = true; };
  $("feuille-ios").onclick = function (e) { if (e.target === this) this.hidden = true; };
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js").catch(function () {});

  // ------------------------------------------------------------ démarrage
  (async function () {
    chargerLocal();
    try { etat.config = await api("/v1/config"); } catch (e) {}
    $("badge-demo").hidden = !etat.config.demo;
    $("form-demo").hidden = !etat.config.demo;
    $("connexion-microsoft").hidden = !!etat.config.demo;
    if (etat.code) {
      try { await ouvrirCompte(); return; } catch (e) { etat.code = ""; sauverLocal(); }
    }
    ecran("connexion");
  })();
})();
