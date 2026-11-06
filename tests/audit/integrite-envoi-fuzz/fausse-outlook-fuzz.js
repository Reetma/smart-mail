// Fausse page « nouvel Outlook » (outlook.cloud.microsoft, interface FR) pour
// l'exploration aléatoire à graines de l'INTÉGRITÉ DE CHAQUE ENVOI de Smart Mail.
//
// Fidèle à la fiche de référence (diagnostic réel) : champ « À » contenteditable
// (aria-label « À », id numérique) avec les pastilles DEDANS (span REK…,
// _EType_RECIPIENT_ENTITY, contenteditable=false, aria-label « offline<NOM> »,
// lpcWrapper / pill / presenceIcon), Cci dans div.LCArn#MSG_<id>_BCC[role=group]
// lui-même dans span.fui-Input avec un <input> vide, étiquettes « À » / « Cci »
// = boutons qui OUVRENT LE CARNET D'ADRESSES (piège), objet aria-label « Objet »
// (placeholder « Ajouter un objet »), corps « Corps du message » (elementToProof,
// Aptos), bouton splitButton-…__primaryActionButton « Envoyer » + « Autres options
// d'envoi », « Réglez votre sensibilité », « Abandonner », « Isoler », « Fermez »,
// ruban avec l'onglet « Message » actif en rédaction et « Importance élevée »
// (aria-pressed), onglet « Modification de <objet> », bandeau, volet, liste des
// messages (role=grid).
//
// Tout ce qui est marqué VARIANTE n'est PAS connu du vrai Outlook : c'est une
// variante plausible, tirée par la graine (window.__CFG, voir generateur.ts).
//
// Ce que la page « envoie » est enregistré dans window.__faux.envoyes (monde
// principal, invisible pour l'extension) ; window.__faux.journal note les clics,
// collages et événements tardifs.
(function () {
  "use strict";
  var CFG = window.__CFG || {};
  var RAPIDE = document.documentElement.dataset.ppTest === "rapide";
  var FAC = RAPIDE ? 0.01 : 1; // même facteur que l'extension (attentes / 100)

  // ------------------------------------------------------------ hasard
  // Une décision = un hachage (graine, clé) : indépendant de l'ordre des événements
  // et des minuteries, donc la même graine donne les mêmes décisions.
  function hacher(s) {
    var h = 2166136261 >>> 0;
    for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
    h ^= h >>> 16; h = Math.imul(h, 0x7feb352d) >>> 0; h ^= h >>> 15; h = Math.imul(h, 0x846ca68b) >>> 0; h ^= h >>> 16;
    return h >>> 0;
  }
  function de(cle) { return hacher(String(CFG.graine) + "|" + cle) / 4294967296; }

  var etat = { redactions: [], active: null, envoyes: [], journal: [], rendus: 0, ouvertures: 0, alertes: [] };
  window.__faux = etat;
  function noter(type, detail) { etat.journal.push({ type: type, detail: detail || null }); }

  var ADR = /^[^\s@;,<>]+@[^\s@;,<>]+\.[^\s@;,<>]{2,}$/;
  var ZW = "​";

  // ------------------------------------------------------------ éléments
  var ROLE = new WeakMap(); // rôle de chaque commande, pour classer les clics (E6)
  function h(tag, attrs, enfants) {
    var el = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === "texte") el.textContent = v;
      else if (k === "roleTest") ROLE.set(el, v);
      else if (k === "style") el.style.cssText = v;
      else el.setAttribute(k, v === true ? "" : v);
    });
    (enfants || []).forEach(function (c) { if (c != null) el.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return el;
  }
  function bouton(libelle, role, extra) {
    var a = { type: "button", class: "fui-Button r1alrhcs", "aria-label": libelle, roleTest: role, texte: libelle };
    Object.keys(extra || {}).forEach(function (k) { a[k] = extra[k]; });
    return h("button", a);
  }
  // Enveloppes DOM supplémentaires (VARIANTE : profondeur inconnue du vrai Outlook).
  function envelopper(el, cle, max) {
    var lim = max != null ? max : (CFG.enveloppes || 0);
    var n = Math.floor(de("env|" + cle) * (lim + 1));
    for (var i = 0; i < n; i++) el = h("div", { class: "___" + hacher(cle + i).toString(36).slice(0, 6) + " f" + i }, [el]);
    return el;
  }
  function profond(el, n, cle) { for (var i = 0; i < (n || 0); i++) el = h("div", { class: "___p" + hacher(cle + i).toString(36).slice(0, 5) }, [el]); return el; }
  function ident(base) { return CFG.idsChangeants ? base + "-" + (++etat.rendus).toString(36) : base; }
  function libelle(el) { return ((el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("title"))) || el.textContent || "").trim().slice(0, 60); }

  // Clics : classés par rôle (autorisé / interdit). Le panneau de l'extension est ignoré.
  document.addEventListener("click", function (e) {
    var t = e.target;
    if (!(t instanceof Element)) return;
    if (t.closest("[data-pp-panneau],[data-testid='ao-panel']")) return;
    var b = t.closest("button,[role=button],[role=tab],[role=menuitem],[role=menuitemradio],[role=option],a[href]");
    if (!b) return;
    var r = ROLE.get(b) || "inconnu:" + libelle(b);
    var red = b.closest("[data-faux-redaction]");
    var d = { role: r, libelle: libelle(b), redaction: red ? red.getAttribute("data-faux-redaction") : null };
    noter(/^(interdit|inconnu)/.test(r) ? "clic-interdit" : "clic", d);
  }, true);

  // ------------------------------------------------------------ annuaire
  var ANNUAIRE = (CFG.annuaire || []).map(function (c) { return { nom: c.nom, adresse: c.adresse }; });
  function contact(adr) {
    var k = String(adr).toLowerCase();
    for (var i = 0; i < ANNUAIRE.length; i++) if (ANNUAIRE[i].adresse.toLowerCase() === k) return ANNUAIRE[i];
    return null;
  }
  // VARIANTE demandée : « suggestion voisine » (jean@x → jean.dupont@x) : même
  // domaine, partie locale qui prolonge celle tapée par « . » ou « - ».
  function voisin(adr) {
    var at = adr.lastIndexOf("@");
    if (at < 1) return null;
    var loc = adr.slice(0, at).toLowerCase(), dom = adr.slice(at + 1).toLowerCase();
    for (var i = 0; i < ANNUAIRE.length; i++) {
      var c = ANNUAIRE[i], a = c.adresse.toLowerCase(), j = a.lastIndexOf("@");
      if (a.slice(j + 1) !== dom) continue;
      var l = a.slice(0, j);
      if (l !== loc && (l.indexOf(loc + ".") === 0 || l.indexOf(loc + "-") === 0)) return c;
    }
    return null;
  }
  function destinataire(adr) {
    var c = contact(adr);
    return { adresse: c ? c.adresse : adr, nom: c ? c.nom : null };
  }
  function affiche(d) {
    if (CFG.affichage === "adresse") return d.adresse;
    return d.nom || d.adresse; // réel : le nom du contact s'il est connu, sinon l'adresse
  }

  // ------------------------------------------------------------ squelette
  var app, ruban, ongletsOuverts, lecture, vueMessage, autreChat;
  function squelette() {
    app = h("div", { id: "app", class: "app" });
    var banniere = h("div", { role: "banner", class: "o365sx" }, [
      h("input", { type: "text", role: "combobox", "aria-label": "Recherchez des e-mails, des réunions, des fichiers et plus encore.", placeholder: "Rechercher", id: "topSearchInput" }),
      bouton("Conversation Teams", "interdit:teams"), bouton("Ma journée", "interdit:ma-journee"),
      bouton("Notifications", "interdit:notifications"), bouton("Paramètres", "interdit:parametres"),
    ]);
    app.appendChild(banniere);
    // VARIANTE : conversation Teams ouverte (volet) avec son propre bouton « Envoyer ».
    if (CFG.autreEnvoyer) {
      autreChat = h("div", { role: "complementary", "aria-label": "Conversation Teams", class: "chat-teams" }, [
        h("div", { texte: "Conversation avec Paul MARTIN" }),
        h("div", { contenteditable: "true", "aria-label": "Tapez un message", role: "textbox", class: "chat-saisie" }),
        bouton("Envoyer", "interdit:teams-envoyer", { title: "Envoyer (Ctrl+Entrée)" }),
      ]);
      app.appendChild(autreChat);
    }
    ruban = h("div", { id: "ruban", class: "ribbon", role: "region", "aria-label": "Ruban" });
    app.appendChild(ruban);
    ongletsOuverts = h("div", { role: "tablist", "aria-label": "Éléments ouverts", class: "onglets-ouverts" });
    app.appendChild(ongletsOuverts);
    var centre = h("div", { class: "centre" });
    centre.appendChild(h("div", { role: "navigation", "aria-label": "Volet de navigation", class: "volet" }, [
      bouton("Courrier", "interdit:volet"), bouton("Calendrier", "interdit:volet"), bouton("Contacts", "interdit:volet"), bouton("To Do", "interdit:volet"),
    ]));
    var grille = h("div", { role: "grid", "aria-label": "Liste de messages", class: "liste" });
    (CFG.listeMessages || []).forEach(function (m, i) {
      grille.appendChild(h("div", { role: "row", "aria-label": m.nom + " " + m.objet, class: "ligne-message", tabindex: "-1" }, [
        h("span", { title: m.adresse, texte: m.nom }), h("span", { texte: m.objet }),
        bouton("Marquer comme non lu", "interdit:liste-non-lu"), bouton("Supprimer", "interdit:liste-supprimer"),
      ]));
    });
    centre.appendChild(grille);
    lecture = h("div", { id: "lecture", role: "main", "aria-label": "Volet de lecture", class: "lecture" });
    var m0 = (CFG.listeMessages || [])[0] || { nom: "Paul MARTIN", adresse: "paul.martin@exemple.fr", objet: "Réunion" };
    vueMessage = h("div", { class: "vue-message" }, [
      h("div", { class: "entete-lu" }, [h("span", { texte: m0.nom }), h("span", { texte: "<" + m0.adresse + ">", title: m0.adresse })]),
      h("div", { texte: "Bonjour, voici le compte rendu de la réunion." }),
      bouton("Répondre", "interdit:repondre"), bouton("Répondre à tous", "interdit:repondre-tous"), bouton("Transférer", "interdit:transferer"),
    ]);
    lecture.appendChild(vueMessage);
    centre.appendChild(lecture);
    app.appendChild(centre);
    document.body.appendChild(app);
    rendreRuban();
    // Doublons cachés présents dès le départ (brouillon de l'utilisateur replié, réponse en ligne).
    (CFG.doublonsCaches || []).forEach(function (dc, i) { if (dc.quand === "toujours") poserDoublon(dc, i, null); });
  }

  // ------------------------------------------------------------ ruban
  function rendreRuban() {
    ruban.textContent = "";
    var c = etat.active;
    var onglets = ["Accueil", "Afficher", "Aide"].concat(c ? ["Message", "Insérer", "Mettre le texte en forme", "Dessiner", "Options"] : []);
    var actif = c ? "Message" : "Accueil";
    var liste = h("div", { role: "tablist", "aria-label": "Onglets du ruban" });
    onglets.forEach(function (o) {
      var t = h("button", { type: "button", role: "tab", "aria-selected": String(o === actif), class: "fui-Tab", roleTest: o === "Message" ? "onglet-message" : "interdit:onglet-" + o, texte: o });
      liste.appendChild(t);
    });
    ruban.appendChild(liste);
    var cmd = h("div", { class: "commandes", role: "toolbar", "aria-label": actif });
    if (!c) {
      var nouveau = CFG.nouveauSansAria
        ? h("button", { type: "button", class: "fui-Button splitPrimary", roleTest: "nouveau", texte: "Nouveau message" }) // VARIANTE : libellé en texte seul
        : h("button", { type: "button", class: "fui-Button splitPrimary", "aria-label": "Nouveau message", roleTest: "nouveau" }, [h("span", { texte: "Nouveau message" })]);
      nouveau.addEventListener("click", ouvrir);
      cmd.appendChild(h("div", { class: "fui-SplitButton" }, [nouveau, bouton("Autres options de nouveau message", "interdit:nouveau-menu")]));
      ["Supprimer", "Archiver", "Signaler", "Déplacer vers", "Répondre", "Répondre à tous", "Transférer"].forEach(function (x) { cmd.appendChild(bouton(x, "interdit:accueil-" + x)); });
    } else {
      ["Coller", "Carnet d'adresses", "Vérifier les noms", "Joindre un fichier", "Lien", "Signature", "Enregistrer"].forEach(function (x) { cmd.appendChild(bouton(x, "interdit:message-" + x)); });
      var imp = h("button", { type: "button", class: "fui-ToggleButton", "aria-label": "Importance élevée", "aria-pressed": String(c.importance === "haute"), roleTest: "importance-haute", texte: "!" });
      imp.addEventListener("click", function () {
        c.importance = c.importance === "haute" ? "normale" : "haute";
        imp.setAttribute("aria-pressed", String(c.importance === "haute"));
        declencher("clic:importance", c);
      });
      var faible = h("button", { type: "button", class: "fui-ToggleButton", "aria-label": "Importance faible", "aria-pressed": String(c.importance === "faible"), roleTest: "interdit:importance-faible", texte: "↓" });
      cmd.appendChild(imp); cmd.appendChild(faible);
      ["Sensibilité", "Chiffrer", "Imprimer un brouillon"].forEach(function (x) { cmd.appendChild(bouton(x, "interdit:message-" + x)); });
      cmd.appendChild(bouton("Autres options", "interdit:ruban-plus", { id: "RibbonOverflowMenu-overflow" }));
    }
    ruban.appendChild(cmd);
  }
  function majOngletOuvert(c) {
    ongletsOuverts.textContent = "";
    if (!c || !c.ouverte) return;
    ongletsOuverts.appendChild(h("div", { role: "tab", "aria-selected": "true", roleTest: "interdit:onglet-modification", texte: "Modification de " + (c.objetValeur() || "(sans objet)") }));
  }

  // ------------------------------------------------------------ rédaction
  var PARA_VIDE = '<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>';
  function blocSignature(html) { return '<div id="Signature" class="elementToProof">' + html + "</div>"; }
  var n = 0;
  function nouvelleRedaction(opts) {
    var c = {
      n: ++n, ouverte: true, cachee: opts.cachee || null, utilisateur: !!opts.utilisateur,
      a: [], cc: [], cci: [], texte: { a: "", cc: "", cci: "" }, enAttente: [],
      ccVisible: !!CFG.ccVisible, cciVisible: !!CFG.cciVisible, importance: "normale",
      objet: opts.objet || "", corpsHTML: "", el: {}, parties: {}, faits: {}, signatures: [], restes: [], pastilles: 0,
    };
    c.objetValeur = function () { return c.el.objet ? c.el.objet.value : c.objet; };
    if (opts.dest) opts.dest.forEach(function (x) { c[x.champ].push(destinataire(x.adresse)); if (x.champ === "cc") c.ccVisible = true; if (x.champ === "cci") c.cciVisible = true; });
    // Corps initial : ligne vide (curseur) puis signature, comme le vrai Outlook.
    var sig = opts.signature !== undefined ? opts.signature : CFG.signature;
    var html = PARA_VIDE;
    if (opts.corps != null) html = opts.corps;
    else if (sig && sig.html && sig.position === "apres") { html = PARA_VIDE + PARA_VIDE + blocSignature(sig.html); c.signatures.push(sig.html); }
    else if (sig && sig.html && sig.position === "avant") { html = blocSignature(sig.html) + PARA_VIDE; c.signatures.push(sig.html); } // VARIANTE
    c.corpsHTML = html;
    return c;
  }

  function construire(c) {
    var racine = h("div", { class: "compose-racine customScrollBar", "data-faux-redaction": String(c.n), role: "region", "aria-label": "Nouveau message" });
    c.racine = racine;
    ["barre", "entete", "objet", "corps"].forEach(function (p) {
      var el = PARTIES[p](c);
      c.parties[p] = el;
      racine.appendChild(el);
    });
    c.dom = profond(envelopper(racine, "racine", CFG.profondeurRacine || 0), 0, "r");
    return c.dom;
  }
  var PARTIES = {
    barre: function (c) {
      var envoyer = h("button", { type: "button", id: ident("splitButton-r" + (200 + c.n) + "__primaryActionButton"), "aria-label": "Envoyer",
        class: "fui-Button r1alrhcs fui-SplitButton__primaryActionButton", roleTest: c.cachee ? "interdit:envoyer-redaction-cachee" : "envoyer", texte: "Envoyer" });
      envoyer.addEventListener("click", function () { clicEnvoyer(c); });
      var split = h("div", { class: "fui-SplitButton" }, [envoyer, bouton("Autres options d'envoi", "interdit:autres-options-envoi", { "aria-haspopup": "menu", texte: "⌄" })]);
      var barre = h("div", { class: "barre-envoi", role: "toolbar", "aria-label": "Commandes de la rédaction" }, [
        split, bouton("Réglez votre sensibilité", "interdit:sensibilite"), bouton("Abandonner", "interdit:abandonner"),
        bouton("Isoler", "interdit:isoler"), bouton("Fermez", "interdit:fermez"),
      ]);
      // VARIANTE : vieux bouton « Envoyer » caché (rédaction précédente, menu replié).
      (CFG.doublonsCaches || []).forEach(function (dc, i) {
        if (dc.quoi !== "envoyer" || c.cachee) return;
        var vieux = h("button", { type: "button", id: "splitButton-r" + (100 + c.n) + "__primaryActionButton", "aria-label": "Envoyer", class: "fui-Button fui-SplitButton__primaryActionButton", roleTest: "interdit:envoyer-perime", texte: "Envoyer" });
        var boite = h("div", { class: "barre-perimee" }, [vieux]);
        cacher(boite, dc.cache);
        if (dc.ordre === "avant") barre.insertBefore(boite, barre.firstChild); else barre.appendChild(boite);
      });
      return envelopper(barre, "barre" + (c.cachee ? "c" : ""));
    },
    entete: function (c) {
      var lignes = [ligneDest(c, "a")];
      if (c.ccVisible) lignes.push(ligneDest(c, "cc"));
      if (c.cciVisible) lignes.push(ligneDest(c, "cci"));
      return envelopper(h("div", { class: "entete-dest" }, lignes), "entete" + (c.cachee ? "c" : ""));
    },
    objet: function (c) {
      var inp = h("input", { type: "text", maxlength: "255", autocomplete: "off", spellcheck: "true", placeholder: "Ajouter un objet", "aria-label": "Objet",
        id: ident("MSG_" + c.n + "_SUBJECT"), class: "fui-Input__input r12stul0" });
      inp.value = c.objet;
      inp.addEventListener("input", function () { c.objet = inp.value; if (c === etat.active) majOngletOuvert(c); saisieCachee(c, "objet"); });
      inp.addEventListener("focusin", function () { declencher("focus:objet", c); });
      inp.addEventListener("focusout", function () { declencher("blur:objet", c); });
      inp.addEventListener("paste", function () { noter("colle", { champ: "objet", n: c.n }); saisieCachee(c, "objet"); });
      c.el.objet = inp;
      return envelopper(profond(h("div", { class: "ligne-objet" }, [h("span", { class: "fui-Input" }, [envelopper(inp, "objet-in" + (c.cachee ? "c" : ""), 2)])]), CFG.profondeurObjet, "po"), "objet" + (c.cachee ? "c" : ""));
    },
    corps: function (c) {
      var d = h("div", { tabindex: "0", dir: "ltr", class: "dFCbN customScrollBar", role: "textbox", "aria-multiline": "true", "aria-label": "Corps du message",
        contenteditable: "true", "data-ms-editor": "true", "data-scrollable": "true", id: ident("editorParent_" + c.n) });
      d.innerHTML = c.corpsHTML;
      d.addEventListener("paste", function (e) { collerCorps(c, d, e); });
      d.addEventListener("input", function () { c.corpsHTML = d.innerHTML; saisieCachee(c, "corps"); });
      d.addEventListener("focusin", function () { declencher("focus:corps", c); });
      d.addEventListener("focusout", function () { declencher("blur:corps", c); });
      c.el.corps = d;
      return envelopper(h("div", { class: "corps-conteneur" }, [d]), "corps" + (c.cachee ? "c" : ""));
    },
  };
  function saisieCachee(c, champ) { if (c.cachee || c.utilisateur) noter("saisie-redaction-cachee", { n: c.n, champ: champ, cachee: c.cachee, utilisateur: c.utilisateur }); }

  var LIB = { a: "À", cc: "Cc", cci: "Cci" }, CODE = { a: "TO", cc: "CC", cci: "BCC" }, ETIQ = { a: "to", cc: "cc", cci: "bcc" }, NUM = { a: "12", cc: "13", cci: "14" };
  function ligneDest(c, champ) {
    var etiquette = h("button", { type: "button", "data-is-focusable": "true", tabindex: "0", role: "button",
      "aria-roledescription": "bouton, ouvre la boîte de dialogue Ajouter des destinataires", class: "fui-Button r1f29ykk HnZn5",
      id: "recipient-well-label-" + ETIQ[champ], roleTest: "interdit:carnet-adresses-" + champ, texte: LIB[champ] });
    etiquette.addEventListener("click", function () { ouvrirCarnet(c); });
    var el = h("div", { "aria-label": LIB[champ], tabindex: "0", class: "___h7sz9k0 f1k6fduh EditorClass", spellcheck: "false", contenteditable: "true", id: ident(NUM[champ]) });
    c.el[champ] = el;
    c[champ].forEach(function (d) { el.appendChild(document.createTextNode(ZW)); el.appendChild(pastille(c, d)); });
    if (c.texte[champ]) el.appendChild(document.createTextNode(c.texte[champ]));
    c.enAttente.forEach(function (j) { if (j.champ === champ && j.texteTemoin) { j.temoin = document.createTextNode(j.texteTemoin); TEMOINS.add(j.temoin); el.appendChild(j.temoin); } });
    brancherChamp(c, champ, el);
    var groupe = h("div", { class: "LCArn", id: ident("MSG_" + c.n + "_" + CODE[champ]), role: "group" }, [el]);
    var span = h("span", { class: "fui-Input r1oeeo9n" }, [envelopper(groupe, "groupe-" + champ + (c.cachee ? "c" : ""), 2), h("input", { type: "text", class: "fui-Input__input", value: "", style: "width:2px;border:0;padding:0" })]);
    var enfants = [etiquette, span];
    if (champ === "a") {
      if (!c.ccVisible) {
        var bcc = h("button", { type: "button", class: "fui-Button bascule", roleTest: "interdit:bascule-cc", texte: "Cc", "aria-label": CFG.basculeCci === "aria-afficher" ? "Afficher Cc" : CFG.basculeCci === "aria" ? "Cc" : null });
        bcc.addEventListener("click", function () { garderTexte(c); c.ccVisible = true; rerendre(c, "entete"); });
        enfants.push(bcc);
      }
      if (!c.cciVisible) {
        var bcci = h("button", { type: "button", class: "fui-Button bascule", roleTest: "bascule-cci", texte: "Cci",
          "aria-label": CFG.basculeCci === "aria-afficher" ? "Afficher Cci" : CFG.basculeCci === "aria" ? "Cci" : null,
          title: CFG.basculeCci === "aria-afficher" ? "Afficher la ligne Cci" : null });
        bcci.addEventListener("click", function () { garderTexte(c); c.cciVisible = true; rerendre(c, "entete"); if (c.el.cci) c.el.cci.focus(); });
        enfants.push(bcci);
      }
    }
    return h("div", { class: "ligne-dest ligne-" + champ }, enfants);
  }

  function pastille(c, d) {
    var nom = affiche(d);
    var s = h("span", { id: CFG.idsChangeants ? "REK" + (++etat.rendus).toString(16).padStart(5, "0") : "REK0000" + (++c.pastilles).toString(16),
      class: "_Entity _EType_RECIPIENT_ENTITY _EReadonly_1 Lbs4W", contenteditable: "false", draggable: "true", "aria-label": "offline" + nom,
      title: CFG.affichage === "nom+titre" ? d.adresse : null }); // VARIANTE nom+titre : l'adresse en info-bulle
    var pill = h("span", { class: d.invalide ? "pill-222 invalidPill-224" : "pill-222 validPill-223", "aria-invalid": d.invalide ? "true" : null }, [h("span", { texte: nom })]);
    var pres = h("span", { id: "presenceIcon", role: "img", "aria-label": "offline", class: "presence" });
    s.appendChild(h("span", { class: "lpcWrapper lpcCommonWeb-hoverTarget" }, [pill, pres]));
    return s;
  }

  // Texte non validé (hors pastilles) d'un champ de destinataires.
  var TEMOINS = new WeakSet();
  function estPastille(nd) { return nd.nodeType === 1 && nd.classList && nd.classList.contains("_Entity"); }
  function texteLibre(el) {
    var s = "";
    [].forEach.call(el.childNodes, function (nd) { if (estPastille(nd) || TEMOINS.has(nd)) return; s += nd.textContent; });
    return s.replace(/​/g, "");
  }
  function viderTexteLibre(el) {
    [].slice.call(el.childNodes).forEach(function (nd) {
      if (estPastille(nd) || TEMOINS.has(nd)) return;
      if (nd.nodeType === 3 && nd.nodeValue === ZW) return;
      nd.remove();
    });
  }
  function garderTexte(c) { ["a", "cc", "cci"].forEach(function (k) { if (c.el[k] && c.el[k].isConnected) c.texte[k] = texteLibre(c.el[k]); }); }

  function brancherChamp(c, champ, el) {
    el.addEventListener("paste", function (e) {
      noter("colle", { champ: champ, n: c.n });
      saisieCachee(c, champ);
      if (CFG.collerDest === "ignore") return; // VARIANTE : collage non pris en charge, l'extension insère elle-même
      var t = e.clipboardData && e.clipboardData.getData("text/plain");
      if (!t) return;
      e.preventDefault();
      var libre = texteLibre(el);
      viderTexteLibre(el);
      saisir(c, champ, libre + t, true);
      declencher("colle:" + champ, c);
    });
    el.addEventListener("input", function () {
      saisieCachee(c, champ);
      var libre = texteLibre(el);
      c.texte[champ] = libre;
      if (!/[;,]/.test(libre)) return;
      viderTexteLibre(el);
      c.texte[champ] = "";
      saisir(c, champ, libre, false);
      if (CFG.collerDest === "ignore") declencher("colle:" + champ, c);
    });
    el.addEventListener("keydown", function (e) { noter("touche", { champ: champ, key: e.key, n: c.n }); });
    el.addEventListener("focusin", function () { declencher("focus:" + champ, c); });
    el.addEventListener("focusout", function () { declencher("blur:" + champ, c); });
  }
  // Découpe « a; b, c » : chaque morceau terminé est résolu en pastille ; au collage,
  // le dernier morceau aussi s'il est complet.
  function saisir(c, champ, texte, collage) {
    var morceaux = texte.split(/[;,]/);
    var reste = morceaux.pop();
    morceaux.forEach(function (m) { resoudre(c, champ, m); });
    reste = (reste || "").replace(/​/g, "");
    if (reste.trim()) {
      if (collage && ADR.test(reste.trim())) resoudre(c, champ, reste);
      else ajouterTexte(c, champ, reste);
    }
  }
  function ajouterTexte(c, champ, t) {
    var el = c.el[champ];
    var nd = document.createTextNode(t);
    if (el && el.isConnected) el.appendChild(nd);
    c.texte[champ] += t;
    return nd;
  }
  function resoudre(c, champ, brut) {
    var tape = String(brut).replace(/[​ ]/g, " ").trim();
    var m = tape.match(/<([^>]+)>/);
    if (m) tape = m[1].trim();
    if (!tape) return;
    var cle = c.n + "|" + champ + "|" + tape.toLowerCase();
    var d = destinataire(tape);
    var invalide = !ADR.test(tape);
    var echec = !!(CFG.echec && de("echec|" + cle) < CFG.echec.proba);
    if (CFG.remplacement && CFG.remplacement.quand === "collage") {
      var v = voisin(tape);
      if (v && de("remplacement|" + cle) < CFG.remplacement.proba) {
        noter("remplacement", { n: c.n, champ: champ, tape: tape, par: v.adresse });
        d = { adresse: v.adresse, nom: v.nom };
      }
    }
    var delai = CFG.delai ? Math.round(CFG.delai.max * de("delai|" + cle)) : 0;
    var job = { fait: false, temoin: null, champ: champ, texteTemoin: null };
    job.faire = function () {
      if (job.fait || !c.ouverte) return;
      job.fait = true;
      if (job.temoin && job.temoin.parentNode) job.temoin.remove();
      c.enAttente = c.enAttente.filter(function (j) { return j !== job; });
      if (echec && CFG.echec.mode === "texte") { noter("echec-resolution", { n: c.n, champ: champ, tape: tape }); ajouterTexte(c, champ, tape); return; }
      var dd = { adresse: d.adresse, nom: d.nom, invalide: invalide || !!(echec && CFG.echec.mode === "invalide") };
      if (dd.invalide) noter("pastille-invalide", { n: c.n, champ: champ, tape: tape });
      ajouterPastille(c, champ, dd);
      if (CFG.doublon && CFG.doublon.quand === "collage" && de("doublon|" + cle) < CFG.doublon.proba) {
        noter("doublon", { n: c.n, champ: champ, adresse: dd.adresse });
        ajouterPastille(c, champ, { adresse: dd.adresse, nom: dd.nom });
      }
    };
    if (!delai) return job.faire();
    // VARIANTE : résolution différée ; le texte reste affiché pendant la résolution, ou disparaît tout de suite.
    if (CFG.delai.mode === "texte-reste") {
      job.texteTemoin = tape + ";";
      job.temoin = document.createTextNode(job.texteTemoin);
      TEMOINS.add(job.temoin);
      if (c.el[champ] && c.el[champ].isConnected) c.el[champ].appendChild(job.temoin);
    }
    c.enAttente.push(job);
    setTimeout(job.faire, delai * FAC);
  }
  function ajouterPastille(c, champ, d) {
    c[champ].push(d);
    var el = c.el[champ];
    if (el && el.isConnected) {
      // la pastille se place avant le texte libre éventuel
      var libre = null;
      for (var i = 0; i < el.childNodes.length; i++) { var nd = el.childNodes[i]; if (!estPastille(nd) && !TEMOINS.has(nd) && !(nd.nodeType === 3 && nd.nodeValue === ZW)) { libre = nd; break; } }
      el.insertBefore(document.createTextNode(ZW), libre);
      el.insertBefore(pastille(c, d), libre);
    }
  }

  // ------------------------------------------------------------ corps
  function texteEnHtml(t) {
    return String(t || "").split(/\r?\n/).map(function (l) { return l ? "<div>" + l.replace(/&/g, "&amp;").replace(/</g, "&lt;") + "</div>" : "<div><br></div>"; }).join("");
  }
  function feuillesTexte(racine) {
    return [].filter.call(racine.querySelectorAll("div,p,li,td,h1,h2,h3,h4,blockquote"), function (b) {
      return !b.querySelector("div,p,li,td,h1,h2,h3,h4,blockquote,ul,ol,table") && (b.textContent || "").replace(/[\s ]/g, "");
    });
  }
  function filtrer(html, c) {
    var t = document.createElement("template");
    t.innerHTML = html;
    var f = CFG.filtreCorps || {};
    if (f.styles) [].forEach.call(t.content.querySelectorAll("[style]"), function (el) { el.removeAttribute("style"); }); // réaliste : styles retirés
    [].slice.call(t.content.querySelectorAll("a[href]")).forEach(function (a, i) {
      if (f.safelinks && de("safelinks|" + c.n + "|" + i) < f.safelinks) {
        noter("safelinks", { n: c.n, href: a.getAttribute("href") });
        a.setAttribute("href", "https://eur02.safelinks.protection.outlook.com/?url=" + encodeURIComponent(a.getAttribute("href")) + "&data=05%7C02%7C%7C&sdata=x&reserved=0");
      } else if (f.retirerLiens && de("delien|" + c.n + "|" + i) < f.retirerLiens) {
        noter("lien-retire", { n: c.n, href: a.getAttribute("href") });
        while (a.firstChild) a.parentNode.insertBefore(a.firstChild, a);
        a.remove();
      }
    });
    if (f.couper === "dernier-bloc") {
      var fe = feuillesTexte(t.content);
      if (fe.length >= 2) { noter("coupe", { n: c.n, quoi: "dernier-bloc", texte: fe[fe.length - 1].textContent.slice(0, 40) }); fe[fe.length - 1].remove(); }
    } else if (f.couper === "listes") {
      [].forEach.call(t.content.querySelectorAll("ul,ol,table"), function (l) { noter("coupe", { n: c.n, quoi: l.tagName }); l.remove(); });
    } else if (f.couper === "apres-n") {
      var total = (t.content.textContent || "").length, garde = Math.floor(total * 0.6), vus = 0, w = document.createTreeWalker(t.content, 4), coupe = false, aRetirer = [];
      for (var nd = w.nextNode(); nd; nd = w.nextNode()) {
        if (coupe) { aRetirer.push(nd); continue; }
        if (vus + nd.nodeValue.length > garde) { nd.nodeValue = nd.nodeValue.slice(0, Math.max(0, garde - vus)); coupe = true; }
        vus += nd.nodeValue.length;
      }
      aRetirer.forEach(function (x) { x.remove(); });
      noter("coupe", { n: c.n, quoi: "apres-n", garde: garde, total: total });
    }
    return t.innerHTML;
  }
  function collerCorps(c, d, e) {
    noter("colle", { champ: "corps", n: c.n });
    saisieCachee(c, "corps");
    var mode = CFG.collerCorps || "html";
    if (mode === "ignore") return; // VARIANTE : l'éditeur ignore le collage synthétique : repli de l'extension (insertHTML)
    e.preventDefault();
    if (mode === "bloque") { noter("collage-bloque", { n: c.n }); return; } // VARIANTE : collage refusé, rien n'entre
    var html = e.clipboardData.getData("text/html"), texte = e.clipboardData.getData("text/plain");
    var a = mode === "texte" || !html ? texteEnHtml(texte) : filtrer(html, c);
    document.execCommand("insertHTML", false, a);
    c.corpsHTML = d.innerHTML;
    declencher("colle:corps", c);
  }

  // ------------------------------------------------------------ événements tardifs (compléments, re-rendus…)
  var ACTIONS = {
    // Complément de signature (VARIANTE plausible) : body.setAsync avec un corps lu à l'ouverture.
    "corps-ecrase": function (c) {
      var sig = (CFG.signatureComplement || "<div>Signature du complément</div>");
      c.el.corps.innerHTML = PARA_VIDE + blocSignature(sig);
      c.signatures = [sig];
      c.corpsHTML = c.el.corps.innerHTML;
    },
    "signature-fin": function (c) { var s = CFG.signatureComplement || "<div>Signature du complément</div>"; c.el.corps.insertAdjacentHTML("beforeend", blocSignature(s)); c.signatures.push(s); },
    "signature-debut": function (c) { var s = CFG.signatureComplement || "<div>Signature du complément</div>"; c.el.corps.insertAdjacentHTML("afterbegin", blocSignature(s)); c.signatures.push(s); },
    // Avant la première ligne vide du corps (c'est là qu'un complément pose la signature
    // d'un message neuf) : après le collage, cette ligne vide est DANS le message.
    "signature-ligne-vide": function (c) {
      var s = CFG.signatureComplement || "<div>Signature du complément</div>";
      var vides = [].filter.call(c.el.corps.querySelectorAll("div"), function (x) { return x.innerHTML === "<br>"; });
      if (vides[0]) vides[0].insertAdjacentHTML("beforebegin", blocSignature(s)); else c.el.corps.insertAdjacentHTML("beforeend", blocSignature(s));
      c.signatures.push(s);
    },
    // setSignatureAsync : remplace le contenu du bloc « Signature » existant.
    "signature-remplace": function (c) {
      var s = CFG.signatureComplement || "<div>Signature du complément</div>";
      var b = c.el.corps.querySelector("#Signature");
      if (b) { b.innerHTML = s; c.signatures = [s]; } else { c.el.corps.insertAdjacentHTML("beforeend", blocSignature(s)); c.signatures.push(s); }
    },
    "doublon-tardif": function (c, t) {
      var champ = t.champ || "a", l = c[champ];
      if (!l.length) return;
      noter("doublon", { n: c.n, champ: champ, adresse: l[l.length - 1].adresse, tardif: true });
      ajouterPastille(c, champ, { adresse: l[l.length - 1].adresse, nom: l[l.length - 1].nom });
    },
    "remplacement-tardif": function (c, t) {
      var champ = t.champ || "a", fait = false;
      c[champ] = c[champ].map(function (d) {
        var v = voisin(d.adresse);
        if (!v || fait) return d;
        fait = true;
        noter("remplacement", { n: c.n, champ: champ, tape: d.adresse, par: v.adresse, tardif: true });
        return { adresse: v.adresse, nom: v.nom };
      });
      if (fait) { garderTexte(c); rerendre(c, "entete"); }
    },
    "importance-reset": function (c) { if (c.importance !== "normale") { c.importance = "normale"; rendreRuban(); } },
    "rerendu-entete": function (c) { garderTexte(c); rerendre(c, "entete", true); },
    "rerendu-objet": function (c) { c.objet = c.el.objet.value; rerendre(c, "objet", true); },
    "rerendu-corps": function (c) { c.corpsHTML = c.el.corps.innerHTML; rerendre(c, "corps", true); },
    "rerendu-tout": function (c) { garderTexte(c); c.objet = c.el.objet.value; c.corpsHTML = c.el.corps.innerHTML; ["barre", "entete", "objet", "corps"].forEach(function (p) { rerendre(c, p, true); }); rendreRuban(); },
    "fenetre-passagere": function () { fenetre("Letsignit Add-in", ["Fermer"], (CFG.dureeFenetre || 3000) * FAC); },
  };
  function declencher(quand, c) {
    if (!c || !c.ouverte) return;
    (CFG.tardifs || []).forEach(function (t, i) {
      if (!(quand === t.quand || (t.quand.slice(-1) === "*" && quand.indexOf(t.quand.slice(0, -1)) === 0))) return;
      if (t.redaction && t.redaction !== c.n) return;
      if (c.cachee) return;
      var cle = i + "|" + c.n;
      if (c.faits[cle]) return;
      c.faits[cle] = 1;
      if (t.proba != null && de("tardif|" + cle) >= t.proba) return;
      noter("tardif", { quoi: t.quoi, quand: quand, n: c.n });
      ACTIONS[t.quoi](c, t);
    });
  }
  // Re-rendu « à la React » d'une partie : nouveaux éléments, même état.
  function rerendre(c, partie, garderFocus) {
    var ancien = c.parties[partie];
    if (!ancien || !ancien.isConnected) return;
    var actif = document.activeElement, quel = null;
    ["a", "cc", "cci", "objet", "corps"].forEach(function (k) { if (c.el[k] && (c.el[k] === actif || (c.el[k].contains && c.el[k].contains(actif)))) quel = k; });
    var neuf = PARTIES[partie](c);
    ancien.replaceWith(neuf);
    c.parties[partie] = neuf;
    noter("rerendu", { n: c.n, partie: partie });
    if (garderFocus && quel && c.el[quel] && CFG.rerenduFocus) {
      try {
        c.el[quel].focus();
        if (quel !== "objet") { var r = document.createRange(); r.selectNodeContents(c.el[quel]); r.collapse(false); var s = getSelection(); s.removeAllRanges(); s.addRange(r); }
      } catch (e) {}
    }
  }

  // ------------------------------------------------------------ doublons cachés (VARIANTES)
  function cacher(el, mode) {
    if (mode === "display") el.style.display = "none";
    else if (mode === "visibility") el.style.visibility = "hidden";
    else if (mode === "hors-ecran") el.style.cssText += ";position:absolute;left:-12000px;top:0;width:640px";
    else if (mode === "opacite") el.style.cssText += ";opacity:0;pointer-events:none";
    else if (mode === "hauteur0") el.style.cssText += ";height:0;overflow:hidden";
    return el;
  }
  function poserDoublon(dc, i, reelle) {
    if (dc.quoi === "redaction-brouillon" || dc.quoi === "redaction-vide") {
      var opts = dc.quoi === "redaction-brouillon"
        ? { cachee: dc.cache, utilisateur: true, objet: "Brouillon perso (ne pas envoyer)", dest: [{ champ: "a", adresse: "paul.martin@exemple.fr" }], corps: "<div>Texte de mon brouillon, à ne pas envoyer.</div>" + PARA_VIDE, signature: null }
        : { cachee: dc.cache, signature: null };
      var c = nouvelleRedaction(opts);
      var el = cacher(h("div", { class: "redaction-repliee" }, [construire(c)]), dc.cache);
      c.conteneur = el;
      etat.redactions.push(c);
      if (reelle) (etat.active.doublons = etat.active.doublons || []).push({ el: el, c: c });
      if (reelle && dc.ordre === "avant") reelle.parentNode.insertBefore(el, reelle);
      else if (reelle) reelle.parentNode.insertBefore(el, reelle.nextSibling);
      else lecture.appendChild(el);
    } else if (dc.quoi === "a-reponse") {
      // Réponse en ligne repliée dans le volet de lecture : son propre champ « À » (pastille de l'expéditeur).
      var r = nouvelleRedaction({ cachee: dc.cache, utilisateur: true, dest: [{ champ: "a", adresse: "paul.martin@exemple.fr" }], signature: null });
      etat.redactions.push(r);
      var ligne = ligneDest(r, "a");
      var bloc = cacher(h("div", { class: "reponse-en-ligne", "data-faux-redaction": String(r.n) }, [ligne]), dc.cache);
      r.conteneur = bloc;
      if (reelle) (etat.active.doublons = etat.active.doublons || []).push({ el: bloc, c: r });
      if (reelle && dc.ordre === "avant") reelle.parentNode.insertBefore(bloc, reelle); else if (reelle) reelle.parentNode.insertBefore(bloc, reelle.nextSibling); else vueMessage.appendChild(bloc);
    }
  }

  // ------------------------------------------------------------ fenêtres
  function fenetre(titre, boutons, duree, bloquante) {
    var id = "Dialog" + (++etat.rendus) + "-title";
    var d = h("div", { role: "dialog", "aria-modal": "true", "aria-labelledby": id, class: "fenetre", style: "position:fixed;left:0;top:0;right:0;bottom:0;z-index:10" }, [
      h("div", { class: "ms-Modal is-open ms-Dialog" }, [h("div", { class: "ms-Overlay ms-Overlay--dark" }), h("div", { class: "ms-Dialog-main", style: "position:fixed;left:35%;top:30%;width:320px;height:160px;background:#fff;border:1px solid #888;z-index:10" }, [
        h("div", { id: id, texte: titre }),
      ].concat(boutons.map(function (b) { return bouton(b, "interdit:fenetre-" + b); })))]),
    ]);
    document.body.appendChild(d);
    if (duree) setTimeout(function () { d.remove(); }, duree);
    return d;
  }
  function alerte(c, texte) {
    noter("alerte", { n: c.n, texte: texte });
    etat.alertes.push(texte);
    var d = h("div", { role: "alertdialog", "aria-modal": "true", class: "alerte", style: "position:fixed;left:30%;top:30%;width:360px;height:140px;background:#fff;border:1px solid #c00;z-index:11" }, [
      h("div", { texte: texte }), bouton("OK", "interdit:alerte-ok"),
    ]);
    document.body.appendChild(d);
  }
  function ouvrirCarnet(c) {
    noter("carnet-ouvert", { n: c.n });
    fenetre("Ajouter des destinataires", ["Enregistrer", "Annuler"], 0);
  }

  // ------------------------------------------------------------ ouvrir / envoyer / fermer
  function ouvrir() {
    etat.ouvertures++;
    if (etat.active && etat.active.ouverte) { noter("nouveau-pendant-redaction", {}); return; }
    var k = etat.ouvertures;
    var opts = {};
    // VARIANTE : rédaction qui garde un reste du message précédent (corps, Cc, Cci).
    var reste = (CFG.restes || []).filter(function (r) { return !r.redaction || r.redaction === k; })[0];
    var c = nouvelleRedaction(opts);
    if (reste) {
      noter("reste", { n: c.n, ou: reste.ou });
      if (reste.ou === "corps") {
        var blocReste = '<div class="elementToProof">' + reste.html + "</div>";
        c.corpsHTML = reste.position === "debut" ? blocReste + c.corpsHTML : c.corpsHTML.replace('<div id="Signature"', blocReste + '<div id="Signature"');
        if (c.corpsHTML.indexOf(blocReste) < 0) c.corpsHTML += blocReste;
        c.restes.push(reste.html);
      } else {
        c[reste.ou].push(destinataire(reste.adresse));
        if (reste.ou === "cc") c.ccVisible = true;
        if (reste.ou === "cci") c.cciVisible = true;
      }
    }
    etat.redactions.push(c);
    etat.active = c;
    vueMessage.style.display = "none";
    var el = construire(c);
    lecture.appendChild(el);
    (CFG.doublonsCaches || []).forEach(function (dc, i) { if (dc.quand === "avec-redaction") poserDoublon(dc, i, el); });
    rendreRuban();
    majOngletOuvert(c);
    // Comme Outlook : le curseur arrive dans « À ».
    try { c.el.a.focus(); } catch (e) {}
    // Complément Letsignit : fenêtre modale passagère à l'ouverture (réelle, durée inconnue).
    if (CFG.fenetreOuverture) fenetre("Letsignit Add-in", ["Fermer"], CFG.fenetreOuverture * FAC);
    // Signature posée par le complément après l'ouverture (VARIANTE : moment exact inconnu).
    declencher("ouverture", c);
  }
  function adresses(l) { return l.map(function (d) { return d.adresse; }); }
  function liensDe(el) {
    return [].map.call(el.querySelectorAll("a[href]"), function (a) { return { href: a.getAttribute("href"), texte: a.textContent }; });
  }
  function clicEnvoyer(c) {
    noter("clic-envoyer", { n: c.n, cachee: c.cachee, utilisateur: c.utilisateur });
    if (!c.ouverte) return;
    if (document.querySelector('[role="alertdialog"]')) return;
    // Résolutions en cours : Outlook les termine avant d'envoyer.
    c.enAttente.slice().forEach(function (j) { j.faire(); });
    // Texte non validé : Outlook le résout au moment d'envoyer (adresse complète) ou refuse.
    ["a", "cc", "cci"].forEach(function (champ) {
      var el = c.el[champ];
      if (!el || !el.isConnected) return;
      var libre = texteLibre(el).replace(/\s+/g, "");
      if (!libre) return;
      libre.split(/[;,]/).filter(Boolean).forEach(function (t) {
        var d = destinataire(t);
        c[champ].push({ adresse: d.adresse, nom: d.nom, invalide: !ADR.test(t) });
      });
      viderTexteLibre(el);
      noter("resolu-a-l-envoi", { n: c.n, champ: champ, texte: libre });
    });
    var tous = c.a.concat(c.cc, c.cci);
    if (!tous.length) return alerte(c, "Ce message doit comporter au moins un destinataire.");
    if (tous.some(function (d) { return d.invalide; })) return alerte(c, "Ce message n'a pas pu être envoyé. Vérifiez les destinataires : une ou plusieurs adresses ne sont pas valides.");
    var corps = c.el.corps;
    etat.envoyes.push({
      n: c.n, cachee: c.cachee, utilisateur: c.utilisateur,
      a: adresses(c.a), cc: adresses(c.cc), cci: adresses(c.cci),
      objet: c.el.objet.value, corpsHTML: corps.innerHTML, corpsTexte: corps.textContent, liens: liensDe(corps),
      importance: c.importance, signatures: c.signatures.slice(), restes: c.restes.slice(),
    });
    fermer(c);
  }
  function fermer(c) {
    c.ouverte = false;
    var el = c.conteneur || c.dom;
    if (el && el.parentNode) el.remove();
    (c.doublons || []).forEach(function (d) { d.c.ouverte = false; if (d.el.parentNode) d.el.remove(); });
    if (etat.active === c) {
      etat.active = null;
      vueMessage.style.display = "";
      rendreRuban();
      majOngletOuvert(null);
    }
  }

  // ------------------------------------------------------------ horloge de test
  // Les étapes affichées par le panneau (« Objet », « Message », « Envoyer »…) servent
  // d'horloge DÉTERMINISTE pour placer un événement tardif (complément, re-rendu)
  // à un instant précis du remplissage : le vrai Outlook peut agir à n'importe quel
  // moment ; on choisit ici lequel, sans minuterie fragile.
  function observerEtapes() {
    var hote = document.querySelector("[data-pp-panneau]");
    var el = hote && hote.shadowRoot && hote.shadowRoot.getElementById("pp-live-etape");
    if (!el) return void setTimeout(observerEtapes, 30);
    var dernier = "";
    new MutationObserver(function () {
      var t = (el.textContent || "").trim();
      if (!t || t === dernier) return;
      dernier = t;
      noter("etape", { texte: t });
      declencher("etape:" + t.split(/[ :,]/)[0], etat.active);
    }).observe(el, { childList: true, characterData: true, subtree: true });
  }

  var css = document.createElement("style");
  css.textContent = "body{font:14px 'Segoe UI',sans-serif;margin:0}.app{display:flex;flex-direction:column;min-height:100vh}.o365sx{display:flex;gap:6px;padding:6px;background:#0f6cbd}" +
    ".ribbon{display:flex;flex-direction:column;gap:2px;padding:4px;border-bottom:1px solid #ddd}.commandes{display:flex;flex-wrap:wrap;gap:3px}.centre{display:flex;gap:8px}.volet{width:90px;display:flex;flex-direction:column}" +
    ".liste{width:260px}.ligne-message{display:block;border-bottom:1px solid #eee;padding:3px}.lecture{flex:1;position:relative;min-width:620px}" +
    ".compose-racine{border:1px solid #ccc;padding:8px;margin:6px}.ligne-dest{display:flex;align-items:center;gap:4px;margin:3px 0}.fui-Input{display:inline-flex;flex:1;align-items:center}" +
    ".EditorClass{min-height:22px;min-width:380px;border-bottom:1px solid #ccc;padding:2px}.r12stul0{width:520px}.dFCbN{min-height:160px;border:1px solid #eee;padding:6px}" +
    "._Entity{background:#eef;border-radius:8px;padding:0 6px;margin:0 2px}.invalidPill-224{color:#c00;text-decoration:underline wavy}.chat-teams{border:1px solid #999;padding:4px;margin:4px;width:300px}" +
    ".chat-saisie{min-height:20px;border:1px solid #ccc}";
  document.head.appendChild(css);
  squelette();
  observerEtapes();
  document.documentElement.dataset.fauxPret = "1";
})();
