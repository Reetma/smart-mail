// Comportement de la fausse page « établissement » (voir outlook-etablissement.html).
//
// window.__politique (posé par le test AVANT le chargement, via addInitScript, et
// modifiable en cours de route par window.__faux.politique({...})) décrit ce que
// l'établissement a configuré. Tout est enregistré dans window.__faux.etat() :
//   envoyes      mails réellement partis (ce que le destinataire recevra)
//   clics        chaque clic (zone : redaction, fenetre, toast, barre, bandeau,
//                liste, recherche, entete, interdit…) avec le nom de la fenêtre
//   saisies      chaque saisie (beforeinput/input/paste) et sa cible
//   fenetres     ouvertures/fermetures de fenêtres
//   boiteEnvoi   mails gardés hors connexion ; brouillons : mails non partis
//   typed        recherches tapées dans « À » (relevé) ; aTexte : texte libre de « À »
// L'état est gardé dans sessionStorage : il survit à un rechargement et à un
// aller-retour vers la page de connexion (même onglet).
(function () {
  "use strict";
  var DOMAINE_INTERNE = "ecole.fr";
  var ANNUAIRE = { "alice@exemple.fr": "MARTIN Alice", "moi@ecole.fr": "TESTEUR Moi", "jean.dupont@ecole.fr": "DUPONT Jean" };
  // Petit carnet pour le relevé : aucune lettre ne remplit la liste (plafond 8),
  // le relevé fait donc exactement les 26 lettres.
  var CARNET = [
    ["Martin Dupont", "martin.dupont@ecole.fr"],
    ["Bruno Roux", "bruno.roux@ecole.fr"],
    ["Claire Petit", "claire.petit@partenaire.com"],
    ["Hugo Vidal", "hugo.vidal@ecole.fr"],
    ["Julie Henry", "julie.henry@ecole.fr"],
    ["Nina Simon", "nina.simon@ecole.fr"],
    ["Paul Robert", "paul.robert@exemple.org"],
    ["Sophie Guerin", "sophie.guerin@ecole.fr"],
  ];
  var PLAFOND = 8;

  // ------------------------------------------------------------ politique et état
  var stockP = {};
  try { stockP = JSON.parse(sessionStorage.getItem("__politique") || "{}"); } catch (e) {}
  var P = (window.__politique = Object.assign({}, window.__politique || {}, stockP));
  var etat;
  try { etat = JSON.parse(sessionStorage.getItem("__etatFaux") || "null"); } catch (e) { etat = null; }
  etat = etat || { envoyes: [], clics: [], saisies: [], fenetres: [], boiteEnvoi: [], brouillons: [], journal: [], typed: [], debuts: [], keysHorsA: [], keysA: [], tentatives: 0, reussis: 0, chargements: 0, clicsEnvoyerSousFenetre: 0 };
  etat.chargements++;
  function sauver() { try { sessionStorage.setItem("__etatFaux", JSON.stringify(etat)); } catch (e) {} }
  function noter(t, d) { etat.journal.push(Object.assign({ t: t, quand: Date.now() }, d || {})); sauver(); }
  sauver();

  function $(id) { return document.getElementById(id); }
  function el(tag, attrs, html) {
    var e = document.createElement(tag);
    for (var k in attrs || {}) { if (attrs[k] != null) e.setAttribute(k, attrs[k]); }
    if (html != null) e.innerHTML = html;
    return e;
  }
  function estPanneau(n) {
    return !!(n && n.closest && n.closest('[data-pp-panneau],[data-testid="ao-panel"]'));
  }
  function decrire(n) {
    if (!n || !n.getAttribute) return "?";
    return n.id || n.getAttribute("aria-label") || n.getAttribute("placeholder") || n.tagName.toLowerCase();
  }
  function zoneDe(n) {
    var z = n && n.closest && n.closest("[data-faux-zone]");
    return z ? z.getAttribute("data-faux-zone") : "autre";
  }
  function fenetreDe(n) {
    var f = n && n.closest && n.closest("[data-fenetre]");
    return f ? f.getAttribute("data-fenetre") : null;
  }

  // ------------------------------------------------------------ journaux de clics et de saisies
  document.addEventListener("click", function (e) {
    var t = e.target;
    if (estPanneau(t)) return;
    var b = t.closest && t.closest('button,[role="button"],[role="menuitem"],[role="option"],[role="tab"],[role="radio"],a,input,[contenteditable="true"]');
    if (!b) return;
    etat.clics.push({ lib: (b.getAttribute("aria-label") || b.textContent || "").trim().slice(0, 60), id: b.id || null, zone: zoneDe(b), fenetre: fenetreDe(b), quand: Date.now() });
    sauver();
  }, true);
  ["beforeinput", "input", "paste"].forEach(function (type) {
    document.addEventListener(type, function (e) {
      var t = e.target;
      if (estPanneau(t)) return;
      var donnees = type === "paste" ? (e.clipboardData && e.clipboardData.getData("text/plain")) : e.data;
      etat.saisies.push({ type: type, cible: decrire(t), zone: zoneDe(t), fenetre: fenetreDe(t), donnees: donnees == null ? null : String(donnees).slice(0, 80) });
      if (etat.saisies.length > 3000) etat.saisies.splice(0, 1000);
      sauver();
    }, true);
  });
  document.addEventListener("keydown", function (e) {
    var t = e.target;
    if (estPanneau(t)) return;
    var a = champ("a");
    if (a && t === a) { etat.keysA.push(e.key); if (etat.keysA.length > 3000) etat.keysA.splice(0, 1000); sauver(); return; }
    etat.keysHorsA.push({ key: e.key, cible: decrire(t), zone: zoneDe(t), fenetre: fenetreDe(t) });
    sauver();
  }, true);

  // ------------------------------------------------------------ pièges à focus (Fluent)
  // Comme FocusTrapZone (Fluent v8) : le focus qui sort d'une fenêtre modale y est
  // ramené aussitôt (synchrone), ou au tour suivant (asynchrone, façon tabster).
  function fenetreAPiege() {
    var ds = document.querySelectorAll("[data-piege]");
    return ds.length ? ds[ds.length - 1] : null;
  }
  document.addEventListener("focusin", function (e) {
    var d = fenetreAPiege();
    if (!d || d.contains(e.target) || estPanneau(e.target)) return;
    var cible = d.querySelector(d.getAttribute("data-piege")) || d;
    if (d.getAttribute("data-piege-mode") === "async") setTimeout(function () { if (d.isConnected) cible.focus(); }, 0);
    else cible.focus();
  }, true);

  // ------------------------------------------------------------ fenêtres, notifications, bandeaux
  var nFenetre = 0;
  // style "v9" (fenêtres d'Outlook : fui-DialogSurface) ou "v8" (compléments :
  // ms-Modal is-open ms-Dialog, comme Letsignit dans le diagnostic réel).
  function ouvrirFenetre(o) {
    nFenetre++;
    var id = o.id || "fenetre" + nFenetre;
    var racine;
    var role = o.alerte ? "alertdialog" : "dialog";
    if (o.style === "v8") {
      racine = el("div", { class: "ms-Layer ms-Layer--fixed", "data-faux-zone": "fenetre", "data-fenetre": id });
      var dlg = el("div", { role: "dialog", "aria-modal": o.modale === false ? null : "true", "aria-labelledby": "Dialog" + nFenetre + "-title", style: "position:fixed;inset:0" });
      var modal = el("div", { class: "ms-Modal is-open ms-Dialog" });
      modal.appendChild(el("div", { class: "ms-Overlay ms-Overlay--dark" }));
      var main = el("div", { class: "ms-Dialog-main" });
      main.appendChild(el("div", { class: "ms-Dialog-title", id: "Dialog" + nFenetre + "-title" }, o.titre || ""));
      if (o.iframe) main.appendChild(el("iframe", { title: o.titre || "Complément", srcdoc: "<p>" + (o.texte || "") + "</p>" }));
      else main.appendChild(el("div", { class: "ms-Dialog-content" }, o.texte || ""));
      racine.appendChild(dlg); dlg.appendChild(modal); modal.appendChild(main);
      var actions = el("div", { class: "ms-Dialog-actions" }); main.appendChild(actions);
      o._actions = actions; o._corps = main; o._dlg = dlg;
    } else if (o.style === "non-modale") {
      racine = el("div", { class: "non-modale", role: "dialog", "aria-label": o.titre || "", "data-faux-zone": "fenetre", "data-fenetre": id });
      racine.appendChild(el("div", { class: "titre" }, "<b>" + (o.titre || "") + "</b>"));
      racine.appendChild(el("div", {}, o.texte || ""));
      var act = el("div", {}); racine.appendChild(act);
      o._actions = act; o._corps = racine; o._dlg = racine;
    } else {
      racine = el("div", { class: "fui-Portal", "data-faux-zone": "fenetre", "data-fenetre": id });
      if (o.modale !== false) racine.appendChild(el("div", { class: "fui-DialogSurface__backdrop", "aria-hidden": "true" }));
      var surf = el("div", { role: role, "aria-modal": o.modale === false ? null : "true", "aria-labelledby": id + "-titre", class: "fui-DialogSurface", tabindex: "-1" });
      var corps = el("div", { class: "fui-DialogBody" });
      corps.appendChild(el("h2", { class: "fui-DialogTitle", id: id + "-titre" }, o.titre || ""));
      corps.appendChild(el("div", { class: "fui-DialogContent" }, o.texte || ""));
      surf.appendChild(corps);
      var act2 = el("div", { class: "fui-DialogActions" }); corps.appendChild(act2);
      racine.appendChild(surf);
      o._actions = act2; o._corps = corps; o._dlg = surf;
    }
    (o.champs || []).forEach(function (c) {
      var f = c.zone ? el("textarea", { "aria-label": c.aria, placeholder: c.placeholder || null, id: c.id || null }) : el("input", { type: "text", "aria-label": c.aria, placeholder: c.placeholder || null, id: c.id || null, value: c.valeur || null });
      o._corps.insertBefore(f, o._actions);
    });
    if (o.listeHtml) o._corps.insertBefore(el("div", {}, o.listeHtml), o._actions);
    (o.boutons || []).forEach(function (b) {
      var x = el("button", { type: "button", "aria-label": b.aria || null, id: b.id || null, class: "fui-Button" }, b.lib);
      x.addEventListener("click", function () {
        noter("bouton-fenetre", { fenetre: id, lib: b.lib });
        if (b.ferme !== false) fermerFenetre(id);
        if (b.action) b.action();
      });
      o._actions.appendChild(x);
    });
    if (o.piege) {
      o._dlg.setAttribute("data-piege", o.piege === true ? "button" : o.piege);
      if (o.piegeMode) o._dlg.setAttribute("data-piege-mode", o.piegeMode);
    }
    var hote = o.dans || document.body;
    hote.appendChild(racine);
    etat.fenetres.push({ id: id, ouverte: Date.now(), titre: o.titre || "" });
    noter("fenetre-ouverte", { fenetre: id });
    fenetresOuvertes[id] = { racine: racine, o: o };
    if (o.piege) { var c0 = o._dlg.querySelector(o._dlg.getAttribute("data-piege")); if (c0) c0.focus(); }
    if (o.duree != null) setTimeout(function () { fermerFenetre(id); }, o.duree);
    return id;
  }
  var fenetresOuvertes = {};
  function fermerFenetre(id) {
    var f = fenetresOuvertes[id];
    if (!f) return;
    delete fenetresOuvertes[id];
    f.racine.remove();
    for (var i = etat.fenetres.length - 1; i >= 0; i--) if (etat.fenetres[i].id === id && !etat.fenetres[i].fermee) { etat.fenetres[i].fermee = Date.now(); break; }
    noter("fenetre-fermee", { fenetre: id });
    if (f.o.auFermer) f.o.auFermer();
  }
  function toast(o) {
    var t = el("div", { class: "fui-Toast", role: o.role || "status", "aria-live": o.role === "alert" ? "assertive" : "polite", "data-faux-zone": "toast", "data-fenetre": o.id || "toast" }, "<div class=\"fui-ToastTitle\">" + o.texte + "</div>");
    var act = el("div", { class: "fui-ToastFooter" }); t.appendChild(act);
    (o.boutons || []).forEach(function (b) {
      var x = el("button", { type: "button", "aria-label": b.aria || null, class: "fui-Button" }, b.lib);
      x.addEventListener("click", function () { noter("bouton-toast", { lib: b.lib }); if (b.action) b.action(); t.remove(); });
      act.appendChild(x);
    });
    $("toaster").appendChild(t);
    noter("toast", { texte: o.texte.replace(/<[^>]+>/g, "").slice(0, 80) });
    if (o.duree != null) setTimeout(function () { t.remove(); }, o.duree);
    return t;
  }
  function barre(hote, o) {
    var b = el("div", { class: "fui-MessageBar", role: o.role || (o.intent === "error" ? "alert" : "status"), "data-intent": o.intent || "info", "data-faux-zone": o.zone || "barre", "data-fenetre": o.id || "barre" },
      "<div class=\"fui-MessageBarBody\">" + o.texte + "</div>");
    var act = el("div", { class: "fui-MessageBarActions" }); b.appendChild(act);
    (o.boutons || []).forEach(function (x) {
      var bt = el("button", { type: "button", "aria-label": x.aria || null, title: x.title || null, class: "fui-Button" }, x.lib);
      bt.addEventListener("click", function () { noter("bouton-barre", { lib: x.lib }); if (x.action) x.action(); });
      act.appendChild(bt);
    });
    hote.appendChild(b);
    noter("barre", { texte: o.texte.replace(/<[^>]+>/g, "").slice(0, 80) });
    return b;
  }
  function bandeau(o) { o.zone = "bandeau"; return barre($("bandeaux"), o); }

  // ------------------------------------------------------------ ruban
  function choisirOnglet(nom) {
    ["accueil", "message"].forEach(function (k) {
      var t = $("onglet-" + k); if (t) t.setAttribute("aria-selected", String(k === nom));
      var pnl = $("ruban-" + k); if (pnl) pnl.hidden = k !== nom;
    });
  }
  $("onglet-accueil").addEventListener("click", function () { choisirOnglet("accueil"); });
  $("onglet-message").addEventListener("click", function () { choisirOnglet("message"); });
  var importance = "normale", etiquette = null;
  $("ruban-importance").addEventListener("click", function () {
    importance = importance === "haute" ? "normale" : "haute";
    this.setAttribute("aria-pressed", String(importance === "haute"));
  });
  $("ruban-carnet").addEventListener("click", function () { ouvrirCarnet(); });
  $("ruban-sensibilite").addEventListener("click", function () { choisirEtiquette(false); });

  // ------------------------------------------------------------ rédaction
  var compteurPastille = 0;
  function champ(nom) {
    var c = $("compose");
    if (!c) return null;
    if (nom === "a") return c.querySelector('[aria-label="À"][contenteditable="true"]');
    if (nom === "cc") return c.querySelector('[aria-label="Cc"][contenteditable="true"]');
    if (nom === "cci") return c.querySelector('[aria-label="Cci"][contenteditable="true"]');
    if (nom === "objet") return c.querySelector('input[aria-label="Objet"]');
    if (nom === "corps") return c.querySelector('[aria-label="Corps du message"]');
    return null;
  }
  function adressesDe(c) { return c ? [].slice.call(c.querySelectorAll("._EType_RECIPIENT_ENTITY")).map(function (p) { return p.dataset.adresse; }) : []; }
  function textesBruts(c) { return [].filter.call(c.childNodes, function (n) { return n.nodeType === 3; }); }
  function texteLibre(c) { return c ? textesBruts(c).map(function (n) { return n.nodeValue; }).join("").replace(/[​﻿]/g, "") : ""; }
  function adresseInvalide(adr) {
    if ((P.invalides || []).map(function (x) { return x.toLowerCase(); }).indexOf(adr.toLowerCase()) >= 0) return true;
    return /\.\.|^\.|\.@/.test(adr);
  }
  function creerPastille(c, adr) {
    var nom = ANNUAIRE[adr.toLowerCase()] || adr;
    var invalide = adresseInvalide(adr);
    c.appendChild(document.createTextNode("​"));
    var p = el("span", { id: "REK0000" + (++compteurPastille).toString(16), class: "_Entity _EType_RECIPIENT_ENTITY _EReadonly_1 Lbs4W", contenteditable: "false", draggable: "true", "aria-label": "offline" + nom });
    p.dataset.adresse = adr;
    var w = el("span", { class: "lpcWrapper" });
    w.appendChild(el("span", { class: invalide ? "pill-222 invalidPill-224" : "pill-222 validPill-223" }, nom.replace(/</g, "&lt;")));
    w.appendChild(el("span", { id: "presenceIcon", role: "img", "aria-label": "offline" }));
    p.appendChild(w);
    c.appendChild(p);
    noter("pastille", { champ: c.getAttribute("aria-label"), adresse: adr, invalide: invalide });
    apresPastille(c, adr);
  }
  function pastilles(c, texte) {
    texte.split(/[;,]/).map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (adr) { creerPastille(c, adr); });
  }
  function brancherDestinataires(c) {
    c.addEventListener("input", function () {
      var t = texteLibre(c);
      if (/[;,]/.test(t)) { textesBruts(c).forEach(function (n) { if (n.nodeValue.replace(/[​﻿]/g, "")) n.remove(); }); pastilles(c, t); cacherSuggestions(); return; }
      if (c.getAttribute("aria-label") === "À") suggestions(c, t.trim());
    });
    c.addEventListener("paste", function (e) {
      if (P.collerIgnoreA) return;
      var t = e.clipboardData && e.clipboardData.getData("text/plain");
      if (!t) return;
      e.preventDefault();
      pastilles(c, t);
    });
  }

  var compose = null, nCompose = 0;
  function ouvrirRedaction(restaurer) {
    if ($("compose")) return;
    nCompose++;
    var z = $("zone");
    z.innerHTML =
      '<div id="compose" class="compose" data-faux-zone="redaction">' +
      '<div class="cmd" role="toolbar" aria-label="Commandes de la rédaction">' +
      '<span class="fui-SplitButton"><button type="button" id="splitButton-r200__primaryActionButton" aria-label="Envoyer" class="fui-Button fui-SplitButton__primaryActionButton">Envoyer</button>' +
      '<button type="button" id="splitButton-r200__menuButton" aria-label="Autres options d\'envoi" class="fui-Button" data-faux-zone="interdit">▾</button></span>' +
      '<button type="button" aria-label="Abandonner" id="abandonner" data-faux-zone="interdit">Abandonner</button>' +
      '<button type="button" aria-label="Réglez votre sensibilité" id="btn-sensibilite" data-faux-zone="interdit">Réglez votre sensibilité</button>' +
      '<button type="button" aria-label="Isoler" data-faux-zone="interdit">Isoler</button>' +
      '<button type="button" aria-label="Fermez" data-faux-zone="interdit">Fermez</button>' +
      "</div>" +
      '<div id="infos"></div>' +
      '<div class="entetes">' +
      '<div class="ligne LCArn" id="MSG_r' + nCompose + '_TO" role="group"><button type="button" id="recipient-well-label-to" aria-roledescription="bouton, ouvre la boîte de dialogue Ajouter des destinataires" class="fui-Button" data-faux-zone="interdit">À</button>' +
      '<div aria-label="À" tabindex="0" class="___h7sz9k0 EditorClass" spellcheck="false" contenteditable="true" id="12"></div>' +
      '<button type="button" id="montrer-cci" aria-label="Afficher le champ Cci" title="Ajouter des destinataires en Cci">Cci</button></div>' +
      '<div class="ligne LCArn" id="MSG_r' + nCompose + '_CC" role="group"><button type="button" id="recipient-well-label-cc" class="fui-Button" data-faux-zone="interdit">Cc</button>' +
      '<div aria-label="Cc" tabindex="0" class="EditorClass" spellcheck="false" contenteditable="true" id="13"></div></div>' +
      '<input type="text" maxlength="255" autocomplete="off" spellcheck="true" placeholder="Ajouter un objet" aria-label="Objet" id="MSG_r' + nCompose + '_SUBJECT" class="fui-Input__input r12stul0">' +
      "</div>" +
      '<div tabindex="0" dir="ltr" class="dFCbN customScrollBar" role="textbox" aria-multiline="true" aria-label="Corps du message" contenteditable="true" data-ms-editor="true" data-scrollable="true">' +
      '<div class="elementToProof" style="font-family: Aptos, sans-serif;"><br></div><div id="Signature"><div>--</div><div>Pierre TESTEUR</div><div>École de test</div></div></div>' +
      "</div>";
    compose = $("compose");
    importance = "normale"; etiquette = null;
    $("ruban-importance").setAttribute("aria-pressed", "false");
    ["message", "inserer", "options"].forEach(function (k) { $("onglet-" + k).hidden = false; });
    choisirOnglet("message");
    brancherDestinataires(champ("a"));
    brancherDestinataires(champ("cc"));
    $("montrer-cci").addEventListener("click", function () {
      var l = el("div", { class: "ligne LCArn", id: "MSG_r" + nCompose + "_BCC", role: "group" },
        '<button type="button" id="recipient-well-label-bcc" data-is-focusable="true" tabindex="0" aria-roledescription="bouton, ouvre la boîte de dialogue Ajouter des destinataires" class="fui-Button r1f29ykk HnZn5" data-faux-zone="interdit">Cci</button>' +
        '<span class="fui-Input"><div aria-label="Cci" tabindex="0" class="EditorClass" spellcheck="false" contenteditable="true" id="14"></div><input type="text" class="fui-Input__input" value="" style="display:none"></span>');
      $("MSG_r" + nCompose + "_CC").after(l);
      this.remove();
      brancherDestinataires(champ("cci"));
      $("recipient-well-label-bcc").addEventListener("click", function () { ouvrirCarnet(); });
    });
    $("recipient-well-label-to").addEventListener("click", function () { ouvrirCarnet(); });
    $("btn-sensibilite").addEventListener("click", function () { choisirEtiquette(false); });
    $("abandonner").addEventListener("click", function () { fermerRedaction("abandon"); });
    $("splitButton-r200__primaryActionButton").addEventListener("click", surEnvoyer);
    var corps = champ("corps");
    corps.addEventListener("paste", function (e) {
      if (P.collerIgnore) return;
      var h = e.clipboardData && e.clipboardData.getData("text/html");
      if (!h) return;
      e.preventDefault();
      document.execCommand("insertHTML", false, h);
      apresCorps();
    });
    var objet = champ("objet");
    objet.addEventListener("paste", function () { apresObjet(); });
    objet.addEventListener("mousedown", function () { surObjetClique(); });
    if (restaurer) {
      restaurer.a.forEach(function (x) { creerPastille(champ("a"), x); });
      objet.value = restaurer.objet;
      corps.innerHTML = restaurer.html;
    }
    // Étiquette de sensibilité : bouton « Envoyer » désactivé tant qu'aucune n'est choisie.
    if (P.etiquette === "desactive") {
      var be = $("splitButton-r200__primaryActionButton");
      be.setAttribute("aria-disabled", "true");
      barre($("infos"), { intent: "warning", id: "etiquette-requise", texte: "Votre organisation exige une étiquette de sensibilité pour envoyer ce message. Choisissez-la avec « Réglez votre sensibilité »." });
    }
    if (P.boitePleine) barre($("infos"), { intent: "warning", id: "boite-pleine-info", texte: "Votre boîte aux lettres est presque pleine." });
    noter("redaction-ouverte", { n: nCompose });
    surOuverture();
    champ("a").focus();
  }
  function fermerRedaction(pourquoi) {
    cacherSuggestions();
    $("zone").innerHTML = "";
    compose = null;
    ["message", "inserer", "options"].forEach(function (k) { $("onglet-" + k).hidden = true; });
    choisirOnglet("accueil");
    noter("redaction-fermee", { pourquoi: pourquoi });
  }
  $("nouveau").addEventListener("click", function () { ouvrirRedaction(null); });

  function lireMessage() {
    return {
      a: adressesDe(champ("a")), cc: adressesDe(champ("cc")), cci: adressesDe(champ("cci")),
      invalides: [].slice.call(compose.querySelectorAll(".invalidPill-224")).map(function (x) { return x.closest("._EType_RECIPIENT_ENTITY").dataset.adresse; }),
      objet: champ("objet").value, corps: champ("corps").innerText, html: champ("corps").innerHTML, importance: importance, etiquette: etiquette,
    };
  }
  function externes(m) { return m.a.concat(m.cc, m.cci).filter(function (x) { return x.split("@").pop().toLowerCase() !== DOMAINE_INTERNE; }); }

  // ------------------------------------------------------------ « Envoyer » : vérifications d'Outlook puis de l'établissement
  function surEnvoyer() {
    var b = $("splitButton-r200__primaryActionButton");
    if (b.disabled || b.getAttribute("aria-disabled") === "true") { noter("envoyer-inactif"); return; }
    etat.tentatives++;
    if (document.querySelector('[data-faux-zone="fenetre"] [aria-modal="true"],[data-faux-zone="fenetre"][aria-modal="true"]')) etat.clicsEnvoyerSousFenetre++;
    sauver();
    var m = lireMessage();
    noter("clic-envoyer", { a: m.a, cci: m.cci, objet: m.objet });
    etapes(m, 0);
  }
  var ETAPES = [
    function aucunDestinataire(m, suite) {
      if (m.a.length || m.cc.length || m.cci.length) return suite();
      ouvrirFenetre({ id: "sans-destinataire", alerte: true, titre: "Impossible d'envoyer", texte: "Ce message doit comporter au moins un destinataire.", boutons: [{ lib: "OK" }] });
    },
    function nomsInvalides(m, suite) {
      if (!m.invalides.length) return suite();
      ouvrirFenetre({ id: "noms", alerte: true, titre: "Ce message n'a pas pu être envoyé", texte: "Vérifiez les noms des destinataires : « " + m.invalides.join(", ") + " » n'est pas une adresse valide.", boutons: [{ lib: "OK" }] });
    },
    function tropDeDestinataires(m, suite) {
      var n = m.a.length + m.cc.length + m.cci.length;
      if (!P.maxDestinataires || n <= P.maxDestinataires) return suite();
      ouvrirFenetre({ id: "trop", alerte: true, titre: "Trop de destinataires", texte: "Ce message compte " + n + " destinataires. Votre organisation en autorise " + P.maxDestinataires + " au maximum par message.", boutons: [{ lib: "OK" }] });
    },
    function sansObjet(m, suite) {
      if (m.objet.trim() || P.sansObjet === false) return suite();
      ouvrirFenetre({ id: "sans-objet", titre: "Ce message n'a pas d'objet", texte: "Voulez-vous l'envoyer quand même ?", boutons: [{ lib: "Envoyer", aria: "Envoyer", action: suite }, { lib: "Ne pas envoyer" }] });
    },
    function pieceJointe(m, suite) {
      if (!P.pieceJointe || !/pi[eè]ce jointe|ci-joint/i.test(m.corps)) return suite();
      ouvrirFenetre({ id: "piece-jointe", titre: "Vous avez oublié une pièce jointe ?", texte: "Votre message mentionne une pièce jointe, mais aucun fichier n'est joint.",
        dans: P.pieceJointe === "dans-redaction" ? compose : null, boutons: [{ lib: "Envoyer", aria: "Envoyer", action: suite }, { lib: "Ne pas envoyer" }] });
    },
    function etiquetteRequise(m, suite) {
      if (!P.etiquette || P.etiquette === "ouverture" || etiquette) return suite();
      if (P.etiquette === "barre") {
        if (!$("etiquette-barre")) barre($("infos"), { intent: "error", id: "etiquette-barre", texte: "Ce message n'a pas été envoyé : sélectionnez une étiquette de sensibilité (« Réglez votre sensibilité »)." });
        return;
      }
      choisirEtiquette(true, suite);
    },
    function destinatairesExternes(m, suite) {
      var ext = externes(m);
      if (!ext.length || (P.externe !== "fenetre" && P.externe !== "fenetre-non-modale")) return suite();
      ouvrirFenetre({ id: "smart-alert", modale: P.externe === "fenetre-non-modale" ? false : true, titre: "Contrôle des destinataires externes",
        texte: "Ce message sera envoyé à des personnes extérieures à l'école (" + ext.slice(0, 3).join(", ") + (ext.length > 3 ? "…" : "") + "). Voulez-vous vraiment l'envoyer ?",
        boutons: [{ lib: "Ne pas envoyer" }, { lib: "Envoyer quand même", action: suite }] });
    },
    function dlp(m, suite) {
      if (P.dlp === "bloquant") {
        return ouvrirFenetre({ id: "dlp", alerte: true, titre: "Message bloqué", texte: "Ce message contient des informations sensibles. La stratégie de protection des données de votre organisation n'autorise pas son envoi.", boutons: [{ lib: "OK" }] });
      }
      if (P.dlp === "justification") {
        return ouvrirFenetre({ id: "dlp-justification", titre: "Ce message contient des informations sensibles", texte: "Pour l'envoyer malgré la stratégie de votre organisation, indiquez une justification.",
          champs: [{ aria: "Justification", zone: true, id: "justification" }], boutons: [{ lib: "Ignorer et envoyer", action: suite }, { lib: "Annuler" }] });
      }
      if (P.dlp === "barre-bloquante") {
        barre($("infos"), { intent: "error", id: "dlp-barre", texte: "Ce message ne peut pas être envoyé : il contient des informations sensibles (stratégie de votre organisation).", boutons: [{ lib: "Ignorer", aria: "Ignorer" }, { lib: "Signaler" }] });
        return;
      }
      suite();
    },
    function boitePleine(m, suite) {
      if (!P.boitePleine) return suite();
      barre($("infos"), { intent: "error", id: "boite-pleine", texte: "Ce message n'a pas pu être envoyé, car votre boîte aux lettres est pleine." });
    },
    function horsLigne(m, suite) {
      if (navigator.onLine) return suite();
      if (P.horsLigne === "refus") {
        barre($("infos"), { intent: "error", id: "hors-ligne", texte: "Impossible d'envoyer ce message : vous êtes hors connexion." });
        return;
      }
      etat.boiteEnvoi.push(m); sauver();
      fermerRedaction("hors-ligne");
      toast({ texte: "Vous êtes hors connexion : le message sera envoyé dès le retour de la connexion (Boîte d'envoi).", duree: 8000 });
    },
    function complementALEnvoi(m, suite) {
      var L = P.letsignit;
      if (!L || L.quand !== "envoi") return suite();
      ouvrirFenetre({ id: "letsignit-envoi", style: "v8", iframe: true, titre: "Letsignit Add-in", texte: "Letsignit vérifie votre signature…", duree: L.duree == null ? null : L.duree,
        auFermer: function () { if (L.duree != null) suite(); } });
    },
    function session(m, suite) {
      if (P.session === "redirection-envoi") { noter("redirection-connexion"); location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=outlook&redirect_uri=https%3A%2F%2Foutlook.office.com%2Fmail%2F"; return; }
      if (P.session === "ecran-envoi") {
        etat.brouillons.push(m); sauver();
        $("app").innerHTML = '<div class="ecran-session" role="main"><h1>Votre session a expiré</h1><p>Actualisez la page pour vous reconnecter.</p><button type="button" data-faux-zone="interdit">Actualiser</button></div>';
        $("ruban").innerHTML = "";
        compose = null;
        noter("ecran-session");
        return;
      }
      suite();
    },
    function echecs(m, suite) {
      var rang = etat.tentatives - 1; // 0 = première tentative d'envoi
      var echec = P.echec && rang >= (P.echecDes || 0) && (P.echecJusqua == null || rang <= P.echecJusqua);
      var limite = P.limitation != null && etat.reussis >= P.limitation;
      if (!echec && !limite) return suite();
      etat.brouillons.push(m); sauver();
      if (P.echec === "reouvre" && !limite) {
        fermerRedaction("envoi");
        setTimeout(function () {
          ouvrirRedaction(m);
          barre($("infos"), { intent: "error", id: "echec-barre", texte: "Ce message n'a pas pu être envoyé. Réessayez plus tard." });
        }, P.delaiEchec == null ? 150 : P.delaiEchec);
        return;
      }
      fermerRedaction("envoi");
      setTimeout(function () {
        toast({ role: "alert", id: "toast-echec", duree: 10000,
          texte: limite ? "Vous avez envoyé trop de messages en peu de temps. Le message « " + m.objet + " » n'a pas été envoyé ; réessayez plus tard."
                        : "Le message « " + m.objet + " » n'a pas pu être envoyé. Il a été enregistré dans le dossier Brouillons.",
          boutons: [{ lib: limite ? "OK" : "Ouvrir le brouillon" }] });
      }, P.delaiEchec == null ? 150 : P.delaiEchec);
    },
    function annulerEnvoi(m, suite) {
      if (!P.annulerEnvoi) return suite();
      fermerRedaction("envoi");
      var reste = Math.ceil(P.annulerEnvoi / 1000), annule = false;
      var t = toast({ id: "annuler-envoi", texte: "Envoi en cours… <span class=\"reste\">" + reste + "</span> s", boutons: [{ lib: "Annuler", aria: "Annuler l'envoi", action: function () { annule = true; noter("envoi-annule"); ouvrirRedaction(m); } }] });
      setTimeout(function () { t.remove(); if (!annule) enregistrerEnvoi(m); }, P.annulerEnvoi);
    },
  ];
  function etapes(m, i) {
    if (i >= ETAPES.length) { enregistrerEnvoi(m); fermerRedaction("envoi"); return; }
    ETAPES[i](m, function () { etapes(lireMessageSiOuvert(m), i + 1); });
  }
  function lireMessageSiOuvert(m) { return compose ? lireMessage() : m; }
  function enregistrerEnvoi(m) {
    etat.envoyes.push({ a: m.a, cc: m.cc, cci: m.cci, objet: m.objet, corps: m.corps, html: m.html, importance: m.importance, etiquette: m.etiquette, quand: Date.now() });
    etat.reussis++;
    sauver();
    noter("envoye", { a: m.a, cci: m.cci.length });
  }
  // Rechargement automatique (mise à jour d'Outlook) peu après le premier clic « Envoyer ».
  document.addEventListener("click", function (e) {
    if (!P.maj || P.maj.rechargerApresEnvoi == null || window.__rechargePrevu) return;
    if (!e.target.closest || !e.target.closest("#splitButton-r200__primaryActionButton")) return;
    if (P.maj.une && etat.chargements > 1) return;
    window.__rechargePrevu = true;
    noter("rechargement-prevu");
    setTimeout(function () { noter("rechargement"); location.reload(); }, P.maj.rechargerApresEnvoi);
  }, false);

  // ------------------------------------------------------------ étiquette de sensibilité
  function choisirEtiquette(aLEnvoi, suite) {
    var liste = ["Public", "Général", "Confidentiel", "Hautement confidentiel"].map(function (x, i) {
      return '<div role="radio" aria-checked="false" tabindex="0" data-etiquette="' + x + '" data-faux-zone="fenetre">' + x + "</div>";
    }).join("");
    ouvrirFenetre({ id: "etiquette", titre: aLEnvoi ? "Étiquette de sensibilité requise" : "Choisir une étiquette de sensibilité",
      texte: "Votre organisation exige qu'une étiquette de sensibilité soit appliquée à ce message.", listeHtml: liste,
      piege: P.etiquettePiege ? '[role="radio"]' : null,
      boutons: [{ lib: "Appliquer", action: function () { etiquette = "Général"; if (aLEnvoi && P.etiquetteEnvoieApresChoix && suite) suite(); } }, { lib: "Annuler" }] });
  }

  // ------------------------------------------------------------ carnet d'adresses (« Ajouter des destinataires »)
  function ouvrirCarnet() {
    var liste = '<div role="listbox" aria-label="Contacts">' + CARNET.map(function (c) {
      return '<div role="option" aria-label="' + c[0] + ", " + c[1] + '">' + c[0] + " <span>" + c[1] + "</span></div>";
    }).join("") + "</div>";
    ouvrirFenetre({ id: "carnet", titre: "Ajouter des destinataires", texte: "", listeHtml: liste,
      champs: [{ aria: "Rechercher des personnes", id: "carnet-recherche" }, { aria: "À", id: "carnet-a" }, { aria: "Cc", id: "carnet-cc" }, { aria: "Cci", id: "carnet-cci" }],
      piege: P.carnetPiege ? "#carnet-recherche" : null,
      boutons: [{ lib: "Enregistrer" }, { lib: "Annuler" }] });
  }

  // ------------------------------------------------------------ compléments (Letsignit et autres)
  function letsignit(o, quand) {
    if (!o || o.quand !== quand) return;
    if (o.fois && (letsignit.vu = (letsignit.vu || 0) + 1) > o.fois) return;
    var id = ouvrirFenetre({ id: "letsignit", style: o.modale === false ? "non-modale" : "v8", modale: o.modale, iframe: o.modale !== false, titre: "Letsignit Add-in",
      texte: "Letsignit insère votre signature…", duree: o.duree == null ? null : o.duree, piege: o.piege || null, piegeMode: o.piegeMode || null,
      champs: o.champ ? [{ aria: o.champ, id: "letsignit-champ" }] : null,
      boutons: [{ lib: "✕", aria: "Fermer" }],
      auFermer: function () { effetComplement(o.effet); } });
    return id;
  }
  function effetComplement(effet) {
    if (!effet || !compose) return;
    if (effet === "reecrit-corps") {
      champ("corps").innerHTML = '<div class="elementToProof"><br></div><div id="Signature"><div>--</div><div>Pierre TESTEUR</div><div>École de test</div><div>Signature Letsignit</div></div>';
      champ("corps").dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertReplacementText" }));
    } else if (effet === "reinsere-signature") {
      var sig = champ("corps").querySelector("#Signature");
      if (!sig) { sig = el("div", { id: "Signature" }); champ("corps").appendChild(sig); }
      sig.innerHTML = "<div>--</div><div>Pierre TESTEUR</div><div>École de test</div>";
    } else if (effet === "vide-objet") {
      champ("objet").value = "";
    } else if (effet === "ajoute-cc") {
      creerPastille(champ("cc"), "archives@ecole.fr");
    }
    noter("effet-complement", { effet: effet });
  }
  var clignotement = null;
  function surOuverture() {
    var L = P.letsignit;
    if (L && L.quand === "ouverture") setTimeout(function () { letsignit(L, "ouverture"); }, 0);
    if (L && L.quand === "clignote") {
      clearInterval(clignotement);
      clignotement = setInterval(function () {
        if (!compose) { clearInterval(clignotement); return; }
        if (!fenetresOuvertes.letsignit) ouvrirFenetre({ id: "letsignit", style: "v8", iframe: true, titre: "Letsignit Add-in", texte: "Letsignit met à jour votre signature…", duree: L.duree || 40 });
      }, L.periode || 90);
    }
    if (P.etiquette === "ouverture") setTimeout(function () { choisirEtiquette(false); }, 0);
    if (P.carnetSurRedaction) setTimeout(function () { ouvrirCarnet(); }, 0);
    if (P.session === "fenetre-ouverture") setTimeout(function () { fenetreSession(); }, 0);
    if (P.maj === "fenetre") setTimeout(function () { fenetreMaj(); }, 0);
    if (P.toastNouveauMessage === "ouverture") setTimeout(function () { toastNouveau(); }, 0);
  }
  var pastilleVue = false;
  function apresPastille(c, adr) {
    if (P.externe === "mailtip" && c.getAttribute("aria-label") === "À" && adr.split("@").pop() !== DOMAINE_INTERNE && !$("mailtip-externe")) {
      barre($("infos"), { intent: "info", id: "mailtip-externe", texte: "Les destinataires suivants sont extérieurs à votre organisation : <span>" + adr + "</span>",
        boutons: [{ lib: "Supprimer les destinataires externes" }, { lib: "✕", aria: "Ignorer" }] });
    }
    if (P.maxDestinataires && !$("mailtip-trop")) {
      var n = adressesDe(champ("a")).length + adressesDe(champ("cc")).length + adressesDe(champ("cci")).length;
      if (n > P.maxDestinataires) barre($("infos"), { intent: "warning", id: "mailtip-trop", texte: "Ce message compte trop de destinataires : votre organisation en autorise " + P.maxDestinataires + " au maximum." });
    }
    if (!pastilleVue && c.getAttribute("aria-label") === "À") {
      pastilleVue = true;
      var L = P.letsignit;
      if (L && L.quand === "apres-a") setTimeout(function () { letsignit(L, "apres-a"); }, 0);
      if (P.toastNouveauMessage === "apres-a") setTimeout(function () { toastNouveau(); }, 0);
    }
  }
  var corpsVu = false;
  function apresCorps() {
    if (P.dlp === "conseil" && !$("dlp-conseil")) {
      barre($("infos"), { intent: "warning", id: "dlp-conseil", texte: "Conseil de stratégie : ce message semble contenir des informations sensibles.", boutons: [{ lib: "Ignorer", aria: "Ignorer" }, { lib: "Signaler un faux positif" }] });
    }
    if (corpsVu) return;
    corpsVu = true;
    var L = P.letsignit;
    if (L && L.quand === "apres-corps") setTimeout(function () { letsignit(L, "apres-corps"); }, 0);
  }
  function apresObjet() {
    if (P.session === "redirection-objet" && etat.chargements === 1) {
      noter("redirection-connexion");
      setTimeout(function () { location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=outlook"; }, 0);
    }
  }
  function surObjetClique() {
    var L = P.letsignit;
    if (L && L.quand === "objet" && !fenetresOuvertes.letsignit) letsignit(L, "objet");
  }
  window.addEventListener("pagehide", function () { pastilleVue = false; corpsVu = false; });

  // ------------------------------------------------------------ session, mise à jour, notifications
  function fenetreSession() {
    ouvrirFenetre({ id: "session", alerte: true, titre: "Votre session a expiré", texte: "Pour continuer, reconnectez-vous.", piege: P.sessionPiege === false ? null : "button",
      boutons: [{ lib: "Se connecter", action: function () { location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=outlook"; } }] });
  }
  function fenetreMaj() {
    ouvrirFenetre({ id: "maj", titre: "Outlook a été mis à jour", texte: "Actualisez la page pour utiliser la nouvelle version.", boutons: [{ lib: "Actualiser", action: function () { location.reload(); } }, { lib: "Plus tard" }] });
  }
  function toastNouveau() {
    toast({ id: "toast-nouveau", texte: "Nouveau message de Jean Dupont (jean.dupont@ecole.fr) — « Réunion demain »", duree: P.dureeToast == null ? 8000 : P.dureeToast,
      boutons: [{ lib: "Répondre" }, { lib: "Marquer comme lu" }, { lib: "Supprimer" }, { lib: "✕", aria: "Fermer" }] });
  }
  function rappel() {
    ouvrirFenetre({ id: "rappel", style: "non-modale", titre: "Rappel : Réunion d'équipe", texte: "Dans 15 minutes — Salle B204",
      boutons: [{ lib: "Rejoindre" }, { lib: "Répéter" }, { lib: "Ignorer", aria: "Ignorer" }, { lib: "Tout ignorer" }] });
  }
  function horsLigneBandeau() {
    if (!navigator.onLine && !document.querySelector('[data-fenetre="bandeau-hors-ligne"]')) bandeau({ intent: "warning", id: "bandeau-hors-ligne", texte: "Vous êtes hors connexion. Certaines fonctionnalités ne sont pas disponibles." });
    if (navigator.onLine) {
      var b = document.querySelector('[data-fenetre="bandeau-hors-ligne"]'); if (b) b.remove();
      // Retour de la connexion : la boîte d'envoi part.
      while (etat.boiteEnvoi.length) enregistrerEnvoi(etat.boiteEnvoi.shift());
    }
  }
  window.addEventListener("offline", horsLigneBandeau);
  window.addEventListener("online", horsLigneBandeau);

  if (P.maj === "bandeau" || (P.maj && P.maj.bandeau)) bandeau({ intent: "info", id: "bandeau-maj", texte: "Une nouvelle version d'Outlook est disponible.", boutons: [{ lib: "Actualiser", title: "Actualiser la page" }, { lib: "Envoyer des commentaires", title: "Envoyer des commentaires" }] });
  if (P.boitePleine) bandeau({ intent: "error", id: "bandeau-boite-pleine", texte: "Votre boîte aux lettres est pleine : vous ne pouvez plus envoyer de messages. Libérez de l'espace.", boutons: [{ lib: "Gérer le stockage" }] });
  if (P.rappel) setTimeout(rappel, 0);
  horsLigneBandeau();

  // ------------------------------------------------------------ suggestions du champ « À » (relevé)
  var seg = "";
  // On note la recherche complète : le plus long texte atteint entre deux retours à vide.
  function track(v) {
    if (v === "") { if (seg) { etat.typed.push(seg); seg = ""; sauver(); declencheurReleve("fin-recherche"); } }
    else { var neuve = !seg; if (v.length >= seg.length) seg = v; if (neuve) { etat.debuts.push(Date.now()); sauver(); declencheurReleve("debut-recherche"); } }
  }
  // Une fenêtre qui s'ouvre PENDANT le relevé, après P.releve.apres recherches :
  //   type "letsignit" (complément, iframe), "champ" (fenêtre avec zone de saisie),
  //   "session" (session expirée : l'annuaire ne répond plus), "toast" (notification).
  var releveFait = false;
  function declencheurReleve(moment) {
    var R = P.releve;
    if (!R || releveFait || (R.moment || "fin-recherche") !== moment || etat.typed.length < (R.apres || 1)) return;
    if (etat.chargements > 1 && !R.toujours) return; // une seule fois (pas de boucle après un rechargement)
    releveFait = true;
    noter("releve-declencheur", { type: R.type, moment: moment, recherches: etat.typed.length });
    if (R.type === "session") {
      P.annuaireEnPanne = true;
      ouvrirFenetre({ id: "session", alerte: true, titre: "Votre session a expiré", texte: "Pour continuer, reconnectez-vous.", piege: R.piege === false ? null : "button", piegeMode: R.piegeMode || null, duree: R.duree == null ? null : R.duree,
        boutons: [{ lib: "Se connecter", action: function () { location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=outlook"; } }] });
    } else if (R.type === "champ") {
      ouvrirFenetre({ id: "copilot", titre: "Brouillon avec Copilot", texte: "Décrivez le message que vous voulez écrire.", champs: [{ aria: "Décrivez ce que vous voulez écrire", id: "copilot-champ" }],
        piege: R.piege === false ? null : "#copilot-champ", piegeMode: R.piegeMode || null, duree: R.duree == null ? null : R.duree, boutons: [{ lib: "Générer" }, { lib: "✕", aria: "Fermer" }] });
    } else if (R.type === "toast") {
      toastNouveau();
    } else if (R.type === "redirection") {
      location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=outlook";
    } else if (R.type === "rechargement") {
      location.reload();
    } else {
      ouvrirFenetre({ id: "letsignit", style: R.modale === false ? "non-modale" : "v8", modale: R.modale, iframe: R.modale !== false, titre: "Letsignit Add-in", texte: "Letsignit met à jour votre signature…",
        piege: R.piege === false ? null : (R.piege || "button"), piegeMode: R.piegeMode || null, duree: R.duree == null ? null : R.duree, boutons: [{ lib: "✕", aria: "Fermer" }] });
    }
    // L'utilisateur demande la pause au même instant (crochet « ao-stop » du panneau, équivalent du bouton « Pause »).
    if (R.pause) document.dispatchEvent(new Event("ao-stop"));
  }
  function tokens(s) {
    var n = String(s || "").toLowerCase().normalize("NFD").replace(/\p{Mn}/gu, "") + " ", t = [], w = "";
    for (var i = 0; i < n.length; i++) { var c = n[i]; if ((c >= "a" && c <= "z") || (c >= "0" && c <= "9")) w += c; else if (w) { t.push(w); w = ""; } }
    return t;
  }
  function chercher(q) {
    var qw = tokens(q);
    return CARNET.filter(function (c) { var tk = tokens(c[0] + " " + c[1]); return qw.every(function (w) { return tk.some(function (t) { return t.indexOf(w) === 0; }); }); }).slice(0, PLAFOND);
  }
  var callout = null;
  function cacherSuggestions() { if (callout) { callout.remove(); callout = null; } }
  function suggestions(c, q) {
    track(q);
    cacherSuggestions();
    if (!q) return;
    var r = c.getBoundingClientRect();
    callout = el("div", { class: "fui-Popover callout", "data-faux-zone": "suggestions", style: "left:" + Math.round(r.left) + "px;top:" + Math.round(r.bottom + 2 + window.scrollY) + "px" });
    var lb = el("div", { role: "listbox", "aria-label": "Suggestions" });
    var horsLigne = !navigator.onLine || P.annuaireEnPanne;
    var res = horsLigne ? [] : chercher(q);
    res.forEach(function (x) {
      var ini = x[0].split(" ").map(function (w) { return w[0]; }).join("");
      lb.appendChild(el("div", { role: "option", "aria-label": x[0] + ", " + x[1] }, '<div class="fui-Persona"><span class="fui-Avatar"><span class="fui-Avatar__initials">' + ini + '</span></span><span class="fui-Persona__primaryText">' + x[0] + '</span><span class="fui-Persona__secondaryText" title="' + x[1] + '">' + x[1] + "</span></div>"));
    });
    callout.appendChild(lb);
    if (horsLigne) callout.appendChild(el("div", { class: "info" }, "Impossible de rechercher dans l'annuaire : vous êtes hors connexion."));
    document.body.appendChild(callout);
  }

  // ------------------------------------------------------------ accès pour les tests
  window.__faux = {
    etat: function () {
      var a = champ("a");
      return Object.assign(JSON.parse(JSON.stringify(etat)), { aTexte: texteLibre(a), redactionOuverte: !!compose, fenetresOuvertes: Object.keys(fenetresOuvertes),
        objetActuel: compose ? champ("objet").value : null, segEnCours: seg });
    },
    politique: function (o) { Object.assign(P, o); try { sessionStorage.setItem("__politique", JSON.stringify(Object.assign(JSON.parse(sessionStorage.getItem("__politique") || "{}"), o))); } catch (e) {} },
    ouvrirFenetre: ouvrirFenetre, fermerFenetre: fermerFenetre, toast: toast, bandeau: bandeau, rappel: rappel, fenetreSession: fenetreSession, fenetreMaj: fenetreMaj,
    toastNouveau: toastNouveau, ouvrirCarnet: ouvrirCarnet, ouvrirRedaction: function () { ouvrirRedaction(null); },
    letsignit: function (o) { return letsignit(Object.assign({ quand: "manuel" }, o), "manuel"); },
  };
  document.documentElement.dataset.fauxPret = "1";
})();
