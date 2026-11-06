// Fausse page Outlook web CONFIGURABLE pour l'audit « Langues et versions ».
// Une seule page, réglée par l'adresse (?lang=de&ui=classique&etroit=1…) :
//
//   lang      fr (défaut) | en | de | es | it | nl | pt | frca | enx   (libelles.js)
//   ui        nouveau (défaut : nouvel Outlook, Fluent v9, fidèle à la fiche de
//             référence) | classique (Outlook web 2018-2023, Fluent v8 :
//             ms-BasePicker, ms-Persona, barre de commandes, « … » > Définir
//             l'importance)
//   theme     clair | sombre (couleurs inversées, attributs data-ogsc posés par
//             l'éditeur après un collage, bouton « Basculer l'arrière-plan »)
//   etroit=1  fenêtre étroite : ruban simplifié, « Importance élevée » rangée dans
//             « Autres options » (menuitemcheckbox + aria-checked)
//   ruban     classique | simplifie (nouveau UI ; défaut : classique, simplifie si étroit)
//   onglet    message (défaut : l'onglet Message s'active à l'ouverture) | accueil
//   fenetre=1 « Nouveau » ouvre la rédaction dans une FENÊTRE SÉPARÉE
//             (window.open de /mail/deeplink/compose)
//   teams=1   volet « Conversation Teams » ouvert (zone de saisie + bouton Envoyer)
//   landmark=0 barre de recherche hors de tout repère (ni role=search, ni banner, ni header)
//   envoiAria libellé (aria-label) du bouton Envoyer, bouton à icône seule
//   impSansEtat=1 importance sans aria-pressed / aria-checked
//   cciVisible=1  ligne Cci affichée d'office (« Toujours afficher Cci ») ; ccVisible=1 : idem pour Cc
//   icone=1   (classique) boutons à icône seule : seulement un title
//   pubs=1    annonce dans la liste des messages (outlook.live.com)
//   dialogue=1 la rédaction s'ouvre dans une fenêtre modale (role=dialog, aria-modal),
//             le reste de l'application est marqué aria-hidden (variante plausible)
//   init / lot  suggestions montrées d'emblée / ajoutées par « plus de résultats »
//
// /mail/deeplink/compose : fenêtre de rédaction séparée (« Isoler », ou fenetre=1).
//
// Tout est journalisé dans window.__journal (clics, saisies, touches, carnet
// d'adresses ouvert, Envoyer de Teams, suggestion choisie) ; les mails « envoyés »
// sont dans data-envoyes (cette fenêtre) et localStorage « __envoyesGlobal » (toutes
// les fenêtres). Aucun réseau.
(function () {
  "use strict";
  var Q = new URLSearchParams(location.search);
  var LANG = Q.get("lang") || "fr";
  var L = window.LIBELLES[LANG] || window.LIBELLES.fr;
  var A = window.ANNUAIRE;
  var UI = Q.get("ui") === "classique" ? "classique" : "nouveau";
  var SOMBRE = Q.get("theme") === "sombre";
  var ETROIT = Q.get("etroit") === "1";
  var RUBAN = Q.get("ruban") || (ETROIT ? "simplifie" : "classique");
  var ONGLET_INITIAL = Q.get("onglet") === "accueil" ? "accueil" : "message";
  var FENETRE = Q.get("fenetre") === "1";
  var TEAMS = Q.get("teams") === "1";
  var LANDMARK = Q.get("landmark") !== "0";
  var ENVOI_ARIA = Q.get("envoiAria");
  var SANS_ETAT = Q.get("impSansEtat") === "1";
  var CCI_VISIBLE = Q.get("cciVisible") === "1";
  var CC_VISIBLE = Q.get("ccVisible") === "1"; // ligne Cc affichée d'office
  var ICONE = Q.get("icone") === "1";
  var PUBS = Q.get("pubs") === "1";
  var DIALOGUE = Q.get("dialogue") === "1"; // rédaction dans une fenêtre modale par-dessus (volet de lecture masqué)
  var INIT = Q.has("init") ? Number(Q.get("init")) : (UI === "classique" ? 2 : 0);
  var LOT = Q.has("lot") ? Number(Q.get("lot")) : 3;
  var POPUP = /\/deeplink\/compose/.test(location.pathname);
  var BROUILLON = Q.get("brouillon");

  document.documentElement.lang = L.lang;
  if (SOMBRE) document.documentElement.classList.add("sombre");
  if (ETROIT) document.documentElement.classList.add("etroit");
  document.documentElement.dataset.ui = UI;

  // ------------------------------------------------------------ journal
  var J = (window.__journal = { clics: [], saisies: [], touches: [], carnet: 0, teamsEnvoye: 0, teamsMessages: [], suggestionChoisie: 0, menus: [], textes: [] });
  window.__razJournal = function () {
    J.clics = []; J.saisies = []; J.touches = []; J.carnet = 0; J.teamsEnvoye = 0; J.teamsMessages = []; J.suggestionChoisie = 0; J.menus = []; J.textes = [];
  };
  function nomChamp(el) { var c = el && el.closest && el.closest("[data-champ]"); return c ? c.getAttribute("data-champ") : (el && el.tagName ? el.tagName.toLowerCase() : "?"); }
  function zoneDe(el) { var z = el && el.closest && el.closest("[data-zone]"); return z ? z.getAttribute("data-zone") : "?"; }
  function horsPanneaux(el) { return !(el && el.closest && el.closest('[data-testid="ao-panel"],[data-pp-panneau]')); }
  document.addEventListener("click", function (e) {
    var t = e.target;
    if (!t || !t.closest || !horsPanneaux(t)) return;
    var b = t.closest('button,[role="button"],[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"],[role="tab"],[role="option"],a,[data-champ]');
    if (!b) return;
    J.clics.push({ lib: (b.getAttribute("aria-label") || b.getAttribute("title") || b.textContent || "").trim().replace(/\s+/g, " ").slice(0, 70),
      id: b.id || "", cmd: b.getAttribute("data-cmd") || "", zone: zoneDe(b), confiance: e.isTrusted });
  }, true);
  document.addEventListener("input", function (e) {
    var t = e.target;
    if (!t || !horsPanneaux(t)) return;
    J.saisies.push({ champ: nomChamp(t), type: e.inputType || "", data: e.data == null ? null : String(e.data), confiance: e.isTrusted });
  }, true);
  document.addEventListener("keydown", function (e) {
    var t = e.target;
    if (!t || !horsPanneaux(t)) return;
    J.touches.push({ champ: nomChamp(t), key: e.key });
    if (e.key === "Escape") fermerMenus();
  }, true);

  // ------------------------------------------------------------ outils
  function h(tag, attrs, enfants) {
    var el = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      var v = attrs[k];
      if (v == null || v === false) continue;
      if (k === "texte") el.textContent = v;
      else if (k === "html") el.innerHTML = v;
      else if (k.indexOf("on") === 0 && typeof v === "function") el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? "" : String(v));
    }
    (enfants || []).forEach(function (c) { if (c) el.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return el;
  }
  // Bouton : libellé visible + aria-label, ou (icone=1, classique) seulement un title.
  function bouton(lib, attrs, contenu) {
    attrs = attrs || {};
    if (ICONE && UI === "classique" && !attrs.garderTexte) { attrs.title = attrs.title || lib; delete attrs["aria-label"]; return h("button", Object.assign({ type: "button" }, attrs), [h("i", { class: "ms-Icon", "aria-hidden": "true", texte: contenu || "" })]); }
    delete attrs.garderTexte;
    return h("button", Object.assign({ type: "button", "aria-label": lib }, attrs), [contenu != null ? contenu : lib]);
  }
  var ZWSP = "​";
  var compteurPastille = 0, compteurRedaction = 0, compteurListe = 0; // pas de hasard : identifiants déterministes
  function hex(n) { return ("0000000" + n.toString(16)).slice(-7); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function initiales(nom) { return String(nom).split(/[\s-]+/).filter(Boolean).slice(0, 2).map(function (m) { return m[0].toUpperCase(); }).join(""); }

  // ------------------------------------------------------------ envois
  window.envoyes = [];
  function publier(mail) {
    window.envoyes.push(mail);
    document.documentElement.dataset.envoyes = JSON.stringify(window.envoyes);
    var g = [];
    try { g = JSON.parse(localStorage.getItem("__envoyesGlobal") || "[]"); } catch (e) { g = []; }
    g.push(mail);
    localStorage.setItem("__envoyesGlobal", JSON.stringify(g));
  }
  document.documentElement.dataset.envoyes = "[]";

  // ------------------------------------------------------------ squelette
  var app = h("div", { id: "app", class: "app" });
  document.body.appendChild(app);
  var lecture, ruban, onglets, commandes, liste, brouillons, barreClassique;
  var courante = null; // rédaction ouverte dans CETTE fenêtre

  function entete() {
    var champRech = h("input", { type: "text", role: "combobox", "aria-expanded": "false", "aria-autocomplete": "list", autocomplete: "off",
      "aria-label": L.recherche, placeholder: L.recherche, "data-champ": "recherche", class: "recherche" });
    var boite = LANDMARK ? h("div", { role: "search", class: "boite-recherche" }, [champRech]) : h("div", { class: "boite-recherche" }, [champRech]);
    var outils = h("div", { class: "outils-haut" }, [
      bouton(L.teams, { "data-cmd": "teams" }, "💬"), bouton("Ma journée", {}, "📅"), bouton("Notifications", {}, "🔔"), bouton("Paramètres", {}, "⚙"),
    ]);
    var kids = [h("div", { class: "logo", texte: "Outlook" }), boite, outils];
    return LANDMARK ? h("header", { role: "banner", class: "entete", "data-zone": "entete" }, kids) : h("div", { class: "entete", "data-zone": "entete" }, kids);
  }

  function construire() {
    if (!POPUP) app.appendChild(entete());
    var milieu = h("div", { class: "milieu" });
    if (!POPUP) milieu.appendChild(h("nav", { class: "rail", "data-zone": "nav", "aria-label": "Applications" }, [
      bouton(L.courrier, {}, "✉"), bouton(L.calendrier, {}, "📆"), bouton(L.contacts, {}, "👥")]));
    var principal = h("div", { class: "principal" });
    milieu.appendChild(principal);
    app.appendChild(milieu);
    if (UI === "nouveau") {
      ruban = h("div", { class: "ruban", "data-zone": "ruban" });
      onglets = h("div", { role: "tablist", class: "onglets", "aria-label": "Ruban" });
      commandes = h("div", { role: "toolbar", class: "commandes", "aria-label": "Commandes" });
      ruban.appendChild(onglets); ruban.appendChild(commandes);
      principal.appendChild(ruban);
      brouillons = h("div", { role: "tablist", class: "brouillons", "aria-label": "Brouillons ouverts" });
      principal.appendChild(brouillons);
    } else if (!POPUP) {
      barreClassique = h("div", { class: "ms-CommandBar barre-app", "data-zone": "ruban" }, [
        bouton(L.nouveauClassique, { class: "ms-Button ms-Button--primary", "data-cmd": "nouveau", onclick: function () { nouveau(); } }, ICONE ? "+" : null),
        bouton("Supprimer", { class: "ms-Button" }, ICONE ? "🗑" : null), bouton("Archiver", { class: "ms-Button" }, ICONE ? "🗄" : null),
      ]);
      principal.appendChild(barreClassique);
    }
    var bas = h("div", { class: "bas" });
    if (!POPUP) {
      liste = h("div", { role: "listbox", "aria-label": L.liste, class: "liste", "data-zone": "liste" });
      var items = [
        ["Apple", L.objetsBoite[0], L.apercu],
        ["Microsoft Learn", L.objetsBoite[1], "Microsoft Learn — learn@mails.microsoft.com"],
        ["Secrétariat", L.objetsBoite[2], "…"],
      ];
      if (PUBS) items.unshift(["Annonce", "Contoso Voyages", "Contoso Voyages · infos@contoso-voyages.example"]);
      items.forEach(function (it, i) {
        liste.appendChild(h("div", { role: "option", "aria-selected": String(i === 0), class: "message-liste", tabindex: "-1" }, [
          h("div", { class: "avatar", texte: initiales(it[0]) }),
          h("div", { class: "texte-liste" }, [h("span", { class: "de", texte: it[0] }), h("span", { class: "sujet", texte: it[1] }), h("span", { class: "apercu", texte: it[2] })]),
          bouton(L.nonLu, { class: "survol" }, "✉"),
        ]));
      });
      bas.appendChild(liste);
    }
    lecture = h("div", { class: "lecture", id: "lecture", "data-zone": "lecture" });
    bas.appendChild(lecture);
    principal.appendChild(bas);
    if (TEAMS && !POPUP) app.appendChild(voletTeams());
  }

  function voletTeams() {
    var saisie = h("div", { role: "textbox", contenteditable: "true", "aria-multiline": "true", "aria-label": L.teamsSaisie, "data-champ": "teams", class: "teams-saisie" });
    return h("aside", { class: "teams", "data-zone": "teams", "aria-label": L.teams }, [
      h("div", { class: "teams-titre", texte: L.teams }),
      h("div", { class: "teams-fil", texte: "Camille : on se voit à 14 h ?" }),
      saisie,
      bouton(L.envoyer, { "data-cmd": "teams-envoyer", class: "teams-envoyer", onclick: function () { J.teamsEnvoye++; var t = saisie.innerText.trim(); if (t) J.teamsMessages.push(t); saisie.textContent = ""; } }, "➤"),
    ]);
  }

  function pageLecture() {
    if (!lecture) return;
    lecture.innerHTML = "";
    if (POPUP) { lecture.appendChild(h("p", { class: "vide", texte: "—" })); return; }
    lecture.appendChild(h("div", { class: "message-ouvert" }, [
      h("h2", { texte: L.objetsBoite[0] }),
      h("div", { class: "actions-lecture" }, [bouton(L.repondre, {}, "↩"), bouton(L.repondreTous, {}, "↩↩"), bouton(L.transferer, {}, "→")]),
      h("p", { texte: L.apercu }),
    ]));
  }

  // ------------------------------------------------------------ ruban (nouveau UI)
  // Comme React/Fluent : les onglets et les commandes ne sont reconstruits que si
  // l'onglet actif change (ou qu'une rédaction s'ouvre/se ferme) ; un clic sur
  // « Importance élevée » met à jour aria-pressed SUR PLACE (mêmes éléments).
  var ongletActif = "accueil", rendu = "";
  function ongletsVoulus() {
    var l = POPUP ? [] : [["accueil", L.accueil], ["afficher", L.afficher], ["aide", L.aide]];
    if (courante) l = l.concat([["message", L.message], ["inserer", L.inserer], ["format", L.format], ["dessiner", L.dessiner], ["options", L.options]]);
    return l;
  }
  function rendreRuban() {
    if (UI !== "nouveau") return;
    var voulus = ongletsVoulus();
    if (!voulus.some(function (o) { return o[0] === ongletActif; })) ongletActif = voulus.length ? voulus[0][0] : "accueil";
    var cle = voulus.map(function (o) { return o[0]; }).join(",") + "|" + ongletActif + "|" + (courante ? courante.id : "");
    if (cle !== rendu) {
      rendu = cle;
      onglets.innerHTML = "";
      voulus.forEach(function (o) {
        var t = h("button", { type: "button", role: "tab", "aria-selected": String(ongletActif === o[0]), "data-onglet": o[0], class: "onglet" }, [o[1]]);
        t.addEventListener("click", function () { ongletActif = o[0]; rendreRuban(); });
        onglets.appendChild(t);
      });
      construireCommandes();
    }
    majImportance();
    rendreBrouillons();
  }
  function majImportance() {
    if (UI !== "nouveau" || SANS_ETAT) return;
    [].forEach.call(commandes.querySelectorAll("[data-imp]"), function (b) { b.setAttribute("aria-pressed", String(!!courante && courante.importance === b.getAttribute("data-imp"))); });
  }
  function construireCommandes() {
    commandes.innerHTML = "";
    if (ongletActif === "accueil") {
      commandes.appendChild(h("span", { class: "fui-SplitButton nouveau" }, [
        bouton(L.nouveau, { class: "fui-Button fui-SplitButton__primaryActionButton primaire", "data-cmd": "nouveau", onclick: function () { nouveau(); } }, "✉ " + L.nouveau),
        bouton(L.nouveauMenu, { class: "fui-Button fui-SplitButton__menuButton", "data-cmd": "nouveau-menu" }, "˅"),
      ]));
      ["Supprimer", "Archiver", "Signaler", "Déplacer vers"].forEach(function (x) { commandes.appendChild(bouton(x, { class: "fui-Button" })); });
    } else if (ongletActif === "message" && courante) {
      var imp = function (lib, niveau) {
        var b = bouton(lib, { class: "fui-ToggleButton", "data-cmd": "imp-" + niveau, "data-imp": niveau }, niveau === "haute" ? "❗" : "⬇");
        b.addEventListener("click", function () { if (!courante) return; courante.importance = courante.importance === niveau ? "normale" : niveau; majImportance(); });
        return b;
      };
      ["Coller", "Gras", "Italique", "Souligner"].forEach(function (x) { commandes.appendChild(bouton(x, { class: "fui-Button" })); });
      if (RUBAN === "classique") {
        commandes.appendChild(bouton(L.carnet, { class: "fui-Button", "data-cmd": "carnet", onclick: function () { carnet(); } }));
        commandes.appendChild(bouton(L.verifier, { class: "fui-Button" }));
      }
      commandes.appendChild(bouton("Joindre un fichier", { class: "fui-Button" }));
      commandes.appendChild(bouton("Lien", { class: "fui-Button" }));
      commandes.appendChild(bouton(L.signature, { class: "fui-Button" }));
      if (RUBAN === "classique") {
        commandes.appendChild(bouton("Enregistrer", { class: "fui-Button" }));
        commandes.appendChild(imp(L.impHaute, "haute"));
        commandes.appendChild(imp(L.impBasse, "faible"));
        ["Sensibilité", "Chiffrer", "Imprimer un brouillon"].forEach(function (x) { commandes.appendChild(bouton(x, { class: "fui-Button" })); });
      }
      var plus = bouton(L.autresOptions, { id: "RibbonOverflowMenu-overflow", class: "fui-Button", "data-cmd": "ruban-plus", "aria-haspopup": "menu" }, "…");
      plus.addEventListener("click", function (e) { e.stopPropagation(); ouvrirMenuRuban(plus); });
      commandes.appendChild(plus);
    } else {
      var autres = { afficher: ["Paramètres d'affichage", "Volet de lecture", "Zoom"], aide: ["Aide", "Commentaires"], inserer: ["Tableau", "Images", "Émoji"],
        format: ["Police", "Taille", "Surligner"], dessiner: ["Stylet", "Gomme"], options: ["Demander un accusé de réception", "Différer la remise", "Afficher De"] }[ongletActif] || ["Règles"];
      autres.forEach(function (x) { commandes.appendChild(bouton(x, { class: "fui-Button" })); });
      // Ruban simplifié / fenêtre étroite : chaque onglet range ses dernières commandes dans « … ».
      if (RUBAN === "simplifie" && courante) {
        var plusO = bouton(L.autresOptions, { class: "fui-Button", "data-cmd": "ruban-plus-" + ongletActif, "aria-haspopup": "menu" }, "…");
        plusO.addEventListener("click", function (e) {
          e.stopPropagation(); fermerMenus();
          var m = h("div", { role: "menu", class: "menu-faux fui-MenuList", "data-zone": "menu" });
          ["Mode Plein écran", "Lecteur immersif"].forEach(function (x) { m.appendChild(itemMenu("menuitem", x)); });
          placerMenu(m, plusO);
        });
        commandes.appendChild(plusO);
      }
    }
  }
  function rendreBrouillons() {
    if (!brouillons) return;
    brouillons.innerHTML = "";
    if (!courante) return;
    var o = courante.champs.objet ? courante.champs.objet.value : "";
    brouillons.appendChild(h("button", { type: "button", role: "tab", "aria-selected": "true", class: "onglet-brouillon" }, [L.modification + " " + (o || L.sansObjet)]));
  }

  // Menus (« Autres options » du ruban, « … » de la rédaction classique).
  function fermerMenus() { [].slice.call(document.querySelectorAll(".menu-faux")).forEach(function (m) { m.remove(); }); }
  document.addEventListener("click", function (e) { if (!e.target.closest || !e.target.closest(".menu-faux")) fermerMenus(); });
  function placerMenu(m, ancre) {
    var r = ancre.getBoundingClientRect();
    m.style.top = Math.round(r.bottom + 2) + "px";
    m.style.left = Math.round(Math.max(0, Math.min(r.left, innerWidth - 260))) + "px";
    document.body.appendChild(m);
  }
  function itemMenu(role, lib, actif, action) {
    var it = h("button", { type: "button", role: role, class: "item-menu" }, [lib]);
    if (actif != null && !SANS_ETAT) it.setAttribute("aria-checked", String(!!actif));
    it.addEventListener("click", function (e) { e.stopPropagation(); J.menus.push(lib); if (action) action(); fermerMenus(); majImportance(); });
    return it;
  }
  function ouvrirMenuRuban(ancre) {
    fermerMenus();
    var c = courante;
    var m = h("div", { role: "menu", class: "menu-faux fui-MenuList", "data-zone": "menu" });
    m.appendChild(itemMenu("menuitem", "Dicter"));
    if (RUBAN === "simplifie" && c) {
      m.appendChild(itemMenu("menuitem", L.carnet, null, function () { carnet(); }));
      m.appendChild(itemMenu("menuitemcheckbox", L.impHaute, c.importance === "haute", function () { c.importance = c.importance === "haute" ? "normale" : "haute"; }));
      m.appendChild(itemMenu("menuitemcheckbox", L.impBasse, c.importance === "faible", function () { c.importance = c.importance === "faible" ? "normale" : "faible"; }));
      m.appendChild(itemMenu("menuitem", "Sensibilité"));
      m.appendChild(itemMenu("menuitem", "Chiffrer"));
    }
    m.appendChild(itemMenu("menuitem", "Lecteur immersif"));
    placerMenu(m, ancre);
  }
  function ouvrirMenuClassique(ancre, c) {
    fermerMenus();
    var m = h("div", { role: "menu", class: "menu-faux ms-ContextualMenu", "data-zone": "menu" });
    m.appendChild(itemMenu("menuitem", "Enregistrer le brouillon"));
    m.appendChild(itemMenu("menuitem", "Insérer une signature"));
    m.appendChild(itemMenu("menuitem", "Afficher De"));
    var sous = h("button", { type: "button", role: "menuitem", class: "item-menu", "aria-haspopup": "menu" }, [L.definirImp]);
    sous.addEventListener("click", function (e) {
      e.stopPropagation();
      J.menus.push(L.definirImp);
      if (m.querySelector(".sous-menu")) return;
      var s = h("div", { role: "menu", class: "sous-menu ms-ContextualMenu" });
      [["haute", L.haute], ["normale", L.normale], ["faible", L.faible]].forEach(function (x) {
        s.appendChild(itemMenu("menuitemradio", x[1], c.importance === x[0], function () { c.importance = x[0]; }));
      });
      m.appendChild(s);
    });
    m.appendChild(sous);
    placerMenu(m, ancre);
  }

  // Carnet d'adresses : fenêtre modale (cliquer le libellé « À »/« Cc »/« Cci » l'ouvre).
  function carnet() {
    J.carnet++;
    var d = h("div", { role: "dialog", "aria-modal": "true", "aria-label": L.carnetTitre, class: "ms-Modal is-open carnet", "data-zone": "dialogue" }, [
      h("h3", { texte: L.carnetTitre }), h("input", { type: "text", "aria-label": L.carnetTitre, "data-champ": "carnet" }),
      bouton(L.ok, { onclick: function () { d.remove(); } }), bouton(L.annuler, { onclick: function () { d.remove(); } }),
    ]);
    document.body.appendChild(d);
  }

  // ------------------------------------------------------------ destinataires
  function estPastille(n) { return n && n.nodeType === 1 && n.matches && n.matches("._EType_RECIPIENT_ENTITY,.ms-PickerPersona-container"); }
  function texteLibreDiv(champ) {
    var out = "";
    var w = document.createTreeWalker(champ, NodeFilter.SHOW_TEXT);
    for (var n = w.nextNode(); n; n = w.nextNode()) {
      if (n.parentElement && n.parentElement.closest("._EType_RECIPIENT_ENTITY")) continue;
      out += n.nodeValue;
    }
    return out.replace(/​/g, "");
  }
  function retirerTexteLibre(champ) {
    var w = document.createTreeWalker(champ, NodeFilter.SHOW_TEXT), aRetirer = [];
    for (var n = w.nextNode(); n; n = w.nextNode()) if (!(n.parentElement && n.parentElement.closest("._EType_RECIPIENT_ENTITY"))) aRetirer.push(n);
    aRetirer.forEach(function (n) { n.remove(); });
    [].slice.call(champ.querySelectorAll("br")).forEach(function (b) { b.remove(); });
  }
  function pastilleNouveau(adr) {
    var nom = A.nomDe(adr);
    var p = h("span", { id: "REK" + hex(++compteurPastille), class: "_Entity _EType_RECIPIENT_ENTITY _EReadonly_1 Lbs4W", contenteditable: "false", draggable: "true", "aria-label": "offline" + nom, "data-adresse": adr });
    p.innerHTML = '<span class="lpcWrapper"><span class="pill-222 validPill-223"><span id="presenceIcon" role="img" aria-label="offline"></span><span class="pillText">' + esc(nom) + "</span></span></span>";
    return p;
  }
  function pastilleClassique(adr, index) {
    var nom = A.nomDe(adr);
    var p = h("div", { class: "ms-PickerPersona-container", role: "listitem", "data-selection-index": String(index), "aria-label": nom, "data-adresse": adr });
    p.innerHTML = '<div class="ms-PickerItem-content"><div class="ms-Persona ms-Persona--size24"><div class="ms-Persona-details"><div class="ms-Persona-primaryText">' + esc(nom) + '</div></div></div></div>';
    p.appendChild(h("button", { type: "button", class: "ms-PickerItem-removeButton", "aria-label": "Supprimer " + nom, tabindex: "-1" }, ["×"]));
    return p;
  }
  function segments(t) { return String(t).split(/[;,\n]/).map(function (x) { return x.trim().replace(/^<|>$/g, ""); }).filter(Boolean); }

  // Callout de suggestions sous un champ destinataire.
  function Suggestions(champ, idListe) {
    var self = this;
    this.champ = champ; this.q = ""; this.montres = INIT;
    if (UI === "nouveau") {
      this.liste = h("div", { role: "listbox", "aria-label": L.suggeres, id: idListe, class: "liste-sugg" });
      this.plus = h("div", { role: "button", tabindex: "-1", class: "plus-resultats", "data-cmd": "plus-resultats" }, ["🔍 " + L.plusResultats]);
      this.callout = h("div", { class: "callout", "data-zone": "suggestions" }, [this.liste, h("div", { class: "pied" }, [this.plus])]);
    } else {
      this.liste = h("div", { role: "listbox", id: idListe, class: "ms-Suggestions-container" });
      this.plus = h("button", { type: "button", class: "ms-SearchMore-button", "data-cmd": "plus-resultats" }, [L.annuaire]);
      this.callout = h("div", { class: "callout ms-Callout", "data-zone": "suggestions" }, [h("div", { class: "ms-Suggestions" }, [h("div", { class: "ms-Suggestions-title", texte: L.suggeres }), this.liste, this.plus])]);
    }
    this.plus.addEventListener("click", function () { self.montres += LOT; self.rendre(); });
  }
  Suggestions.prototype.rendre = function () {
    var self = this, tous = this.q ? A.ask(this.q) : [];
    var vus = tous.slice(0, this.montres);
    this.liste.innerHTML = "";
    vus.forEach(function (c, i) {
      var o;
      if (UI === "nouveau") {
        o = h("div", { role: "option", id: self.liste.id + "-o" + i, "aria-selected": String(i === 0), "aria-label": c.name + ", " + c.mail, class: "option-sugg" });
        o.innerHTML = '<div class="fui-Persona"><span class="fui-Avatar"><span class="fui-Avatar__initials">' + esc(initiales(c.name)) + '</span></span><span class="fui-Persona__primaryText">' + esc(c.name) +
          '</span><span class="fui-Persona__secondaryText" title="' + esc(c.mail) + '">' + esc(c.mail) + "</span></div>";
      } else {
        o = h("div", { role: "option", id: self.liste.id + "-o" + i, class: "ms-Suggestions-item" });
        o.innerHTML = '<button type="button" class="ms-Suggestions-itemButton" tabindex="-1"><div class="ms-Persona"><div class="ms-Persona-details"><div class="ms-Persona-primaryText">' + esc(c.name) +
          '</div><div class="ms-Persona-secondaryText">' + esc(c.mail) + "</div></div></div></button>";
      }
      o.addEventListener("click", function (e) { e.preventDefault(); J.suggestionChoisie++; self.surChoix(c.mail); });
      self.liste.appendChild(o);
    });
    this.plus.hidden = !(this.q && tous.length > this.montres);
    var ouvert = !!this.q;
    if (ouvert && !this.callout.isConnected) document.body.appendChild(this.callout);
    if (!ouvert && this.callout.isConnected) this.callout.remove();
    if (ouvert) {
      var r = this.champ.getBoundingClientRect();
      this.callout.style.top = Math.round(r.bottom + 4) + "px";
      this.callout.style.left = Math.round(r.left) + "px";
    }
    if (UI === "classique") this.champ.setAttribute("aria-expanded", String(ouvert));
  };
  Suggestions.prototype.chercher = function (q) {
    q = String(q || "").trim();
    if (q !== this.q) { this.q = q; this.montres = INIT; }
    this.rendre();
  };
  Suggestions.prototype.premiere = function () { var t = this.q ? A.ask(this.q).slice(0, Math.max(1, this.montres)) : []; return t[0] ? t[0].mail : null; };

  function brancherNouveau(champ, nom) {
    var sugg = new Suggestions(champ, "sugg-" + nom + "-" + champ.id);
    function ajouter(adresses) {
      retirerTexteLibre(champ);
      adresses.forEach(function (adr) { champ.appendChild(pastilleNouveau(adr)); champ.appendChild(document.createTextNode(ZWSP)); });
      sugg.chercher("");
    }
    sugg.surChoix = function (mail) { ajouter([mail]); };
    champ.addEventListener("input", function () {
      var libre = texteLibreDiv(champ);
      J.textes.push({ champ: nom, libre: libre });
      if (/[;,]/.test(libre)) { ajouter(segments(libre)); return; }
      sugg.chercher(libre);
    });
    champ.addEventListener("paste", function (e) {
      var t = e.clipboardData && e.clipboardData.getData("text/plain");
      if (!t) return;
      e.preventDefault();
      ajouter(segments(t));
      apresCollage(nom);
    });
    champ.addEventListener("keydown", function (e) {
      if (e.key !== "Enter" && e.key !== "Tab") return;
      var libre = texteLibreDiv(champ).trim();
      if (!libre) return;
      e.preventDefault();
      ajouter([sugg.premiere() || libre]); // Entrée/Tab valide la suggestion en tête (comme Outlook)
    });
    champ.addEventListener("blur", function () { setTimeout(function () { if (document.activeElement !== champ) sugg.chercher(""); }, 150); });
    return sugg;
  }
  function brancherClassique(input, zoneListe, nom) {
    var sugg = new Suggestions(input, "sugg-" + nom + "-" + (++compteurListe));
    input.setAttribute("aria-controls", sugg.liste.id);
    function ajouter(adresses) {
      adresses.forEach(function (adr) { zoneListe.appendChild(pastilleClassique(adr, zoneListe.children.length)); });
      input.value = "";
      sugg.chercher("");
    }
    sugg.surChoix = function (mail) { ajouter([mail]); };
    input.addEventListener("input", function () {
      J.textes.push({ champ: nom, libre: input.value });
      if (/[;,]/.test(input.value)) { ajouter(segments(input.value)); return; }
      sugg.chercher(input.value);
    });
    input.addEventListener("keydown", function (e) {
      if (e.key !== "Enter" && e.key !== "Tab") return;
      var v = input.value.trim();
      if (!v) return;
      e.preventDefault();
      ajouter([sugg.premiere() || v]);
    });
    input.addEventListener("blur", function () { setTimeout(function () { if (document.activeElement !== input) sugg.chercher(""); }, 150); });
    return sugg;
  }

  // ------------------------------------------------------------ rédaction
  // « L'utilisateur clique Isoler » à un moment PRÉCIS (déterministe) : dès que
  // l'extension a collé dans tel champ (synchrone), ou dès qu'elle survole le bouton
  // Envoyer pour cliquer dessus (__isolerAuSurvolEnvoyer).
  function apresCollage(nomChampColle) {
    if (window.__isolerApres === nomChampColle && courante) isoler(courante);
  }
  function ligneDestNouveau(c, cle, libelle, avecBoutons) {
    var idCh = String(10 + 2 * compteurRedaction + (cle === "a" ? 2 : cle === "cc" ? 3 : 4));
    var lab = h("button", { type: "button", "data-is-focusable": "true", tabindex: "0", role: "button", "aria-roledescription": L.carnetRole, class: "fui-Button r1f29ykk HnZn5 etiquette", id: "recipient-well-label-" + (cle === "a" ? "to" : cle === "cc" ? "cc" : "bcc"), "data-cmd": "carnet-" + cle }, [libelle]);
    lab.addEventListener("click", function () { carnet(); });
    var champ = h("div", { "aria-label": libelle, tabindex: "0", class: "___h7sz9k0 EditorClass champ-dest", spellcheck: "false", contenteditable: "true", id: idCh, "data-champ": cle });
    var enfants = [lab];
    if (cle === "a") enfants.push(champ);
    else enfants.push(h("span", { class: "fui-Input" }, [champ, h("input", { type: "text", class: "fui-Input__input cache-fluent", value: "" })]));
    if (avecBoutons) {
      if (!CC_VISIBLE) {
        var bCc = h("button", { type: "button", class: "fui-Button montrer", "aria-label": L.montrerCc, title: L.montrerCc, "data-cmd": "montrer-cc" }, [L.cc]);
        bCc.addEventListener("click", function () { montrerLigne(c, "cc"); bCc.remove(); });
        enfants.push(bCc);
      }
      if (!CCI_VISIBLE) {
        var bCci = h("button", { type: "button", class: "fui-Button montrer", "aria-label": L.montrerCci, title: L.montrerCci, "data-cmd": "montrer-cci" }, [L.cci]);
        bCci.addEventListener("click", function () { montrerLigne(c, "cci"); bCci.remove(); });
        enfants.push(bCci);
      }
    }
    var ligne = h("div", { class: "LCArn ligne-dest", id: "MSG_" + c.id + "_" + (cle === "a" ? "TO" : cle.toUpperCase()), role: "group" }, enfants);
    c.champs[cle] = champ;
    c.sugg[cle] = brancherNouveau(champ, cle);
    return ligne;
  }
  function ligneDestClassique(c, cle, libelle, avecBoutons) {
    var zoneListe = h("div", { class: "ms-SelectionZone", role: "list" });
    var input = h("input", { class: "ms-BasePicker-input", "aria-label": libelle, role: "combobox", "aria-autocomplete": "both", "aria-expanded": "false", autocomplete: "off", "data-champ": cle });
    var lab = bouton(libelle, { class: "ms-Button etiquette", title: L.carnetTitre, "data-cmd": "carnet-" + cle, garderTexte: true }, libelle);
    lab.addEventListener("click", function () { carnet(); });
    var enfants = [lab, h("div", { class: "ms-BasePicker" }, [h("div", { class: "ms-BasePicker-text" }, [zoneListe, input])])];
    if (avecBoutons && !CCI_VISIBLE) {
      var bCci = h("button", { type: "button", class: "ms-Link montrer", "aria-label": L.montrerCci, title: L.montrerCci, "data-cmd": "montrer-cci" }, [L.cci]);
      bCci.addEventListener("click", function () { montrerLigne(c, "cci"); bCci.remove(); });
      enfants.push(bCci);
    }
    var ligne = h("div", { class: "ligne-dest classique" }, enfants);
    c.champs[cle] = input; c.listes[cle] = zoneListe;
    c.sugg[cle] = brancherClassique(input, zoneListe, cle);
    return ligne;
  }
  function montrerLigne(c, cle) {
    if (c.champs[cle]) return;
    var lib = cle === "cc" ? L.cc : L.cci;
    var l = UI === "nouveau" ? ligneDestNouveau(c, cle, lib, false) : ligneDestClassique(c, cle, lib, false);
    var apres = cle === "cci" && c.lignes.cc ? c.lignes.cc : c.lignes.a;
    apres.after(l);
    c.lignes[cle] = l;
  }

  function nouvelleRedaction(init) {
    init = init || {};
    var c = { id: (POPUP ? "p" : "m") + (++compteurRedaction), importance: init.importance || "normale", champs: {}, lignes: {}, listes: {}, sugg: {} };
    var zone = "redaction-" + c.id;
    var objet = h("input", { type: "text", maxlength: "255", autocomplete: "off", spellcheck: "true", placeholder: L.objetPh, "aria-label": UI === "nouveau" ? L.objet : L.objet + ",", id: "MSG_" + c.id + "_SUBJECT", class: UI === "nouveau" ? "fui-Input__input r12stul0 objet" : "ms-TextField-field objet", "data-champ": "objet" });
    objet.addEventListener("input", function () { rendreBrouillons(); if (window.__isolerApres === "objet" && courante === c) isoler(c); });
    var corps = h("div", { tabindex: "0", dir: "ltr", class: UI === "nouveau" ? "dFCbN customScrollBar corps" : "corps classique", role: "textbox", "aria-multiline": "true", "aria-label": L.corps, contenteditable: "true", "data-ms-editor": "true", "data-scrollable": "true", "data-champ": "corps" });
    corps.innerHTML = init.corpsHtml != null ? init.corpsHtml : '<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt;"><br></div>' +
      '<div id="Signature"><div>-- </div><div>Équipe Smart Mail</div><div>Paris</div></div>';
    corps.addEventListener("paste", function (e) {
      var html = e.clipboardData && e.clipboardData.getData("text/html");
      if (!html) return;
      e.preventDefault();
      document.execCommand("insertHTML", false, html);
      if (SOMBRE) [].forEach.call(corps.querySelectorAll("p,div,span,a,b,strong,li"), function (el) { if (!el.hasAttribute("data-ogsc")) { el.setAttribute("data-ogsc", "rgb(0, 0, 0)"); el.style.color = "rgb(255, 255, 255)"; } });
      apresCollage("corps");
    });
    var tete, lignes = [];
    if (UI === "nouveau") {
      var envoyerB = ENVOI_ARIA
        ? h("button", { type: "button", id: "splitButton-r200__primaryActionButton", "aria-label": ENVOI_ARIA, class: "fui-Button fui-SplitButton__primaryActionButton envoyer", "data-cmd": "envoyer" }, [h("span", { class: "fui-Button__icon", "aria-hidden": "true", texte: "➤" })])
        : h("button", { type: "button", id: "splitButton-r200__primaryActionButton", "aria-label": L.envoyer, class: "fui-Button fui-SplitButton__primaryActionButton envoyer", "data-cmd": "envoyer" }, [L.envoyer]);
      envoyerB.addEventListener("click", function () { envoyer(c); });
      envoyerB.addEventListener("pointerover", function () { if (window.__isolerAuSurvolEnvoyer && courante === c) isoler(c); });
      var plusEnvoi = h("button", { type: "button", "aria-label": L.envoyerPlus, class: "fui-Button fui-SplitButton__menuButton", "data-cmd": "envoyer-plus", "aria-haspopup": "menu" }, ["˅"]);
      var outils = [h("span", { class: "fui-SplitButton" }, [envoyerB, plusEnvoi]), bouton(L.sensibilite, { class: "fui-Button" }, "🏷")];
      if (SOMBRE) outils.push(bouton(L.arrierePlan, { class: "fui-Button", "data-cmd": "arriere-plan", onclick: function () { corps.classList.toggle("fond-clair"); } }, "☀"));
      outils.push(bouton(L.abandonner, { class: "fui-Button", "data-cmd": "abandonner", onclick: function () { fermerRedaction(c); } }, "🗑"));
      if (!POPUP) outils.push(bouton(L.isoler, { class: "fui-Button", "data-cmd": "isoler", onclick: function () { isoler(c); } }, "⧉"));
      outils.push(bouton(L.fermer, { class: "fui-Button", "data-cmd": "fermer", onclick: function () { fermerRedaction(c); } }, "✕"));
      tete = h("div", { class: "tete-redaction", role: "toolbar", "aria-label": "Rédaction" }, outils);
      c.lignes.a = ligneDestNouveau(c, "a", L.a, true);
      lignes.push(c.lignes.a);
    } else {
      var envoyerC = bouton(L.envoyer, { class: "ms-Button ms-CommandBarItem-link envoyer", title: L.envoyer + " (" + L.raccourciEnvoi + ")", "data-cmd": "envoyer" }, ICONE ? "➤" : null);
      envoyerC.addEventListener("click", function () { envoyer(c); });
      var plusC = bouton(L.plusClassique, { class: "ms-Button ms-CommandBar-overflowButton", "data-cmd": "redaction-plus", "aria-haspopup": "true" }, "…");
      plusC.addEventListener("click", function (e) { e.stopPropagation(); ouvrirMenuClassique(plusC, c); });
      tete = h("div", { class: "ms-CommandBar tete-redaction", role: "menubar" }, [envoyerC, bouton(L.joindre, { class: "ms-Button" }, ICONE ? "📎" : null),
        bouton(L.ignorerClassique, { class: "ms-Button", "data-cmd": "abandonner", onclick: function () { fermerRedaction(c); } }, ICONE ? "🗑" : null), plusC]);
      c.lignes.a = ligneDestClassique(c, "a", L.a, true);
      lignes.push(c.lignes.a);
    }
    var racine = h("div", { class: "redaction " + UI, "data-zone": zone, "data-redaction": c.id }, [tete].concat(lignes).concat([h("div", { class: "ligne-objet" }, [UI === "nouveau" ? h("span", { class: "fui-Input" }, [objet]) : h("div", { class: "ms-TextField" }, [objet])]), corps]));
    c.el = racine; c.champs.objet = objet; c.champs.corps = corps;
    if (CCI_VISIBLE) { lignes[0].after(UI === "nouveau" ? ligneDestNouveau(c, "cci", L.cci, false) : ligneDestClassique(c, "cci", L.cci, false)); c.lignes.cci = lignes[0].nextSibling; }
    if (CC_VISIBLE) { lignes[0].after(UI === "nouveau" ? ligneDestNouveau(c, "cc", L.cc, false) : ligneDestClassique(c, "cc", L.cc, false)); c.lignes.cc = lignes[0].nextSibling; }
    return c;
  }

  function adressesDe(c, cle) {
    var ch = c.champs[cle];
    if (!ch) return [];
    var pills = UI === "nouveau" ? [].slice.call(ch.querySelectorAll("._EType_RECIPIENT_ENTITY")) : [].slice.call(c.listes[cle].querySelectorAll(".ms-PickerPersona-container"));
    var out = pills.map(function (p) { return p.getAttribute("data-adresse"); });
    var libre = UI === "nouveau" ? texteLibreDiv(ch).trim() : ch.value.trim();
    if (libre) out.push("TEXTE-LIBRE:" + libre); // Outlook tenterait de le résoudre : destinataire en trop
    return out;
  }
  function etatRedaction(c) {
    if (!c) return null;
    var f = function (cle) { var ch = c.champs[cle]; if (!ch) return null; return { adresses: adressesDe(c, cle), libre: UI === "nouveau" ? texteLibreDiv(ch) : ch.value, pastilles: UI === "nouveau" ? ch.querySelectorAll("._EType_RECIPIENT_ENTITY").length : c.listes[cle].children.length }; };
    return { id: c.id, a: f("a"), cc: f("cc"), cci: f("cci"), objet: c.champs.objet.value, corpsTexte: c.champs.corps.innerText, corpsHtml: c.champs.corps.innerHTML, importance: c.importance };
  }

  function envoyer(c) {
    if (courante !== c) return;
    var a = adressesDe(c, "a"), cc = adressesDe(c, "cc"), cci = adressesDe(c, "cci");
    if (!a.length && !cc.length && !cci.length) return; // Outlook refuse : aucun destinataire
    publier({ a: a, cc: cc, cci: cci, objet: c.champs.objet.value, corps: c.champs.corps.innerText, html: c.champs.corps.innerHTML, importance: c.importance, fenetre: POPUP ? "separee" : "principale" });
    fermerRedaction(c);
  }
  function fermerRedaction(c) {
    if (courante !== c) return;
    ["a", "cc", "cci"].forEach(function (k) { if (c.sugg[k]) c.sugg[k].chercher(""); });
    c.el.remove();
    if (c.dialogue) { c.dialogue.remove(); app.removeAttribute("aria-hidden"); }
    courante = null;
    ongletActif = "accueil";
    pageLecture();
    rendreRuban();
    if (POPUP) document.title = "Fenêtre fermée";
  }
  function ouvrirRedaction(init) {
    if (courante) return courante;
    var c = nouvelleRedaction(init);
    courante = c;
    if (DIALOGUE && !POPUP) {
      c.dialogue = h("div", { role: "dialog", "aria-modal": "true", "aria-label": L.nouveauClassique, class: "ms-Modal is-open modale-redaction", "data-zone": "redaction-" + c.id }, [c.el]);
      document.body.appendChild(c.dialogue);
      app.setAttribute("aria-hidden", "true");
    } else {
      lecture.innerHTML = "";
      lecture.appendChild(c.el);
    }
    ongletActif = UI === "nouveau" ? (ONGLET_INITIAL === "accueil" ? "accueil" : "message") : ongletActif;
    rendreRuban();
    // Comme Outlook : le curseur va dans « À ».
    setTimeout(function () { if (courante === c) c.champs.a.focus(); }, 0);
    return c;
  }
  function nouveau() {
    if (FENETRE && !POPUP) {
      var qs = new URLSearchParams(location.search); qs.delete("fenetre");
      window.open("/mail/deeplink/compose?" + qs.toString(), "redaction" + Date.now(), "popup,width=1100,height=760");
      return;
    }
    ouvrirRedaction();
  }
  // « Isoler » : la rédaction part dans une fenêtre séparée (même contenu), et
  // disparaît de la fenêtre principale.
  function isoler(c) {
    if (courante !== c) return;
    var st = { a: adressesDe(c, "a"), cc: adressesDe(c, "cc"), cci: adressesDe(c, "cci"), objet: c.champs.objet.value, corpsHtml: c.champs.corps.innerHTML, importance: c.importance };
    localStorage.setItem("brouillon-" + c.id, JSON.stringify(st));
    var qs = new URLSearchParams(location.search); qs.delete("fenetre"); qs.set("brouillon", c.id);
    window.open("/mail/deeplink/compose?" + qs.toString(), "brouillon" + c.id, "popup,width=1100,height=760");
    ["a", "cc", "cci"].forEach(function (k) { if (c.sugg[k]) c.sugg[k].chercher(""); });
    c.el.remove();
    if (c.dialogue) { c.dialogue.remove(); app.removeAttribute("aria-hidden"); }
    courante = null;
    ongletActif = "accueil";
    pageLecture();
    rendreRuban();
  }
  function remplir(c, st) {
    ["a", "cc", "cci"].forEach(function (k) {
      var l = (st[k] || []).filter(function (x) { return x.indexOf("TEXTE-LIBRE:") !== 0; });
      var libres = (st[k] || []).filter(function (x) { return x.indexOf("TEXTE-LIBRE:") === 0; }).map(function (x) { return x.slice(12); });
      if (!l.length && !libres.length) return;
      if (!c.champs[k]) montrerLigne(c, k);
      var ch = c.champs[k];
      l.forEach(function (adr) {
        if (UI === "nouveau") { ch.appendChild(pastilleNouveau(adr)); ch.appendChild(document.createTextNode(ZWSP)); }
        else c.listes[k].appendChild(pastilleClassique(adr, c.listes[k].children.length));
      });
      if (libres.length) { if (UI === "nouveau") ch.appendChild(document.createTextNode(libres.join(" "))); else ch.value = libres.join(" "); }
    });
    if (st.objet) c.champs.objet.value = st.objet;
    if (st.corpsHtml != null) c.champs.corps.innerHTML = st.corpsHtml;
    if (st.importance) c.importance = st.importance;
    rendreRuban();
  }

  // ------------------------------------------------------------ crochets de test
  // Brouillon que L'UTILISATEUR a déjà commencé : destinataires, objet, texte, et
  // l'endroit où se trouve son curseur.
  window.__ouvrirBrouillon = function (o) {
    o = o || {};
    var c = ouvrirRedaction();
    remplir(c, { a: o.a || [], cc: o.cc || [], objet: o.objet || "", corpsHtml: o.corps != null ? "<div>" + esc(o.corps) + "</div>" + '<div id="Signature"><div>-- </div><div>Équipe Smart Mail</div></div>' : null });
    return new Promise(function (r) { setTimeout(function () { window.__placerCurseur(o.focus || null, o.caret || "fin"); r(true); }, 20); });
  };
  window.__placerCurseur = function (cle, caret) {
    var el = cle === "recherche" ? document.querySelector('[data-champ="recherche"]') : cle && courante ? courante.champs[cle] : null;
    if (!el) { if (document.activeElement && document.activeElement.blur) document.activeElement.blur(); return false; }
    el.focus();
    if (el.tagName === "INPUT") { var n = caret === "debut" ? 0 : el.value.length; el.setSelectionRange(n, n); }
    else {
      // caret "texte" (corps) : à la fin du texte de l'utilisateur, juste avant la signature.
      var cible = caret === "texte" && el.firstElementChild ? el.firstElementChild : el;
      var r = document.createRange(); r.selectNodeContents(cible); r.collapse(caret === "debut"); var s = getSelection(); s.removeAllRanges(); s.addRange(r);
    }
    return true;
  };
  window.__isoler = function () { if (courante) isoler(courante); };
  window.__etat = function () {
    return { redaction: etatRedaction(courante), dialogue: !!document.querySelector('.carnet[role="dialog"]'), menus: document.querySelectorAll(".menu-faux").length,
      onglet: ongletActif, actif: nomChamp(document.activeElement), recherche: (document.querySelector('[data-champ="recherche"]') || {}).value || "",
      teams: (document.querySelector('[data-champ="teams"]') || {}).textContent || "" };
  };
  window.__nouveau = nouveau;

  // ------------------------------------------------------------ démarrage
  construire();
  if (POPUP) {
    var st = null;
    try { st = BROUILLON ? JSON.parse(localStorage.getItem("brouillon-" + BROUILLON) || "null") : null; } catch (e) { st = null; }
    document.title = (st ? L.modification + " " + (st.objet || L.sansObjet) : L.nouveau) + " - Outlook";
    var cP = ouvrirRedaction();
    if (st) remplir(cP, st);
  } else {
    pageLecture();
    rendreRuban();
  }
  document.documentElement.dataset.fauxPret = "1";
})();
