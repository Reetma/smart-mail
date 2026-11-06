/* Fausse page FIDÈLE au nouvel Outlook web : outlook.cloud.microsoft, interface FR,
 * établissement avec le complément de signature Letsignit.
 *
 * CONNU (fiche de référence, diagnostic réel) et reproduit tel quel :
 *  - champ « À » : <div aria-label="À" tabindex="0" class="… EditorClass" spellcheck="false"
 *    contenteditable="true" id="12"> (pas de role, pas d'aria-controls) ; Cc et Cci pareils,
 *    dans <div class="LCArn" id="MSG_<id>_BCC" role="group">, lui-même dans un span.fui-Input
 *    qui contient AUSSI un <input class="fui-Input__input" value=""> vide ;
 *  - libellés « À », « Cc », « Cci » = BOUTONS qui ouvrent le carnet d'adresses
 *    « Ajouter des destinataires » (id recipient-well-label-to|cc|bcc) : le piège ;
 *  - pastille validée DANS le champ : ​<span id="REK…" class="_Entity _EType_RECIPIENT_ENTITY
 *    _EReadonly_1 Lbs4W" contenteditable="false" draggable="true" aria-label="offline<NOM>">
 *    avec span.lpcWrapper, span.pill-222.validPill-223, #presenceIcon : elle montre le NOM ;
 *  - Objet : input aria-label « Objet », placeholder « Ajouter un objet » ;
 *  - corps : div role=textbox aria-label « Corps du message » contenteditable, paragraphes
 *    div.elementToProof (Aptos) ; la signature du complément est insérée dans le corps ;
 *  - Envoyer scindé (splitButton-…__primaryActionButton aria-label « Envoyer ») + « Autres
 *    options d'envoi » ; « Réglez votre sensibilité », « Abandonner », « Isoler », « Fermez » ;
 *  - ruban : Accueil, Afficher, Aide, Message (actif en rédaction), Insérer, Mettre le texte
 *    en forme, Dessiner, Options ; dans Message … « Importance élevée » (aria-pressed
 *    false→true), « Importance faible », …, « Autres options » (RibbonOverflowMenu-overflow) ;
 *  - barre de recherche combobox « Recherchez des e-mails, des réunions… », Teams, Ma journée,
 *    Notifications, Paramètres ; volet de gauche ; liste des messages (listbox « Liste des
 *    messages », aperçus avec adresses) ;
 *  - fenêtre « Letsignit Add-in » : role=dialog aria-modal + .ms-Modal.is-open + ms-Overlay--dark ;
 *  - suggestions du champ À en Fluent v9 (spans collés, adresse répétée dans title et
 *    aria-label, badge « Externe »), ligne « Afficher plus de résultats » hors de la listbox.
 *
 * NON CONNU, modélisé en variantes PLAUSIBLES réglables (window.__OUTLOOK_CFG) : délais,
 * suggestions au focus, conversion en pastille au départ du champ, « ; » qui valide la
 * suggestion, collage ignoré, ruban étroit, fenêtres de politique, hors connexion, etc.
 *
 * Tout ce qui se passe est noté dans window.__journal (clics, touches, saisies, collages,
 * focus, pastilles, fenêtres) ; les mails « envoyés » dans <html data-envoyes>.
 */
(function () {
  "use strict";
  var H = document.documentElement;

  var DEFAUTS = {
    rapide: true,              // <html data-pp-test="rapide"> : attentes de l'extension / 100
    echelle: null,             // facteur des délais d'Outlook (null : 0,01 si rapide, sinon 1)
    annuaire: null,            // [[nom, adresse, externe?, poids?], …]
    initial: 3,                // contacts montrés d'emblée pour une recherche
    lot: 5,                    // contacts ajoutés par « Afficher plus de résultats »
    suggestionsAuFocus: true,  // « Contacts suggérés » dès qu'on entre dans un champ vide
    initialFocus: 5,
    delaiSuggestions: 150,     // ms avant mise à jour de la liste après une frappe
    delaiResolution: 300,      // ms : la pastille passe de l'adresse au nom (annuaire)
    collageIgnore: [],         // champs qui ignorent un collage synthétique : "A", "CC", "CCI", "CORPS"
    pointVirgule: "texte",     // « ; » arrivé par une saisie : "texte" (ce qui est tapé) | "suggestion" (meilleure suggestion)
    flou: "pastille",          // texte laissé dans un champ de destinataires quand on le quitte : "pastille" | "texte"
    cciVisible: false,         // réglage « Toujours afficher Cci »
    ongletRedaction: "Message",// onglet du ruban sélectionné à l'ouverture d'une rédaction
    rubanEtroit: false,        // fenêtre étroite : « Importance élevée » rangée dans « Autres options »
    importanceDelai: 120,      // ms avant que aria-pressed suive le clic
    letsignit: null,           // { html, delai, absolu, conteneur:"Signature"|null, modale:{ apres, duree, piege } }
    envoiDialogue: null,       // "sensibilite" | "externes" : fenêtre de politique à l'envoi
    horsLigne: "file",         // envoi hors connexion : "file" (Boîte d'envoi, partira au retour) | "echec"
    brouillon: null,           // { a:[adresses], objet, corps, focus:"A"|"CORPS"|"OBJET"|null } ouvert au chargement
    lecture: 1,                // message affiché dans le volet de lecture (index) ou null
    reponse: null,             // { depuis:index, objetVisible:false, focus:"CORPS" } ouverte au chargement
    reponseObjetVisible: false,
    messages: null,
    carnetChamps: false,       // le carnet d'adresses contient des champs « À »/« Cci » éditables
    moi: "moi@ecole-test.fr",
  };
  var CFG = Object.assign({}, DEFAUTS, window.__OUTLOOK_CFG || {});
  window.__cfg = CFG;
  if (CFG.rapide === false) H.removeAttribute("data-pp-test"); else H.setAttribute("data-pp-test", "rapide");
  var ECH = CFG.echelle != null ? CFG.echelle : (CFG.rapide === false ? 1 : 0.01);
  function D(ms, absolu) { return Math.max(0, Math.round(absolu ? ms : ms * ECH)); }
  function echapper(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  var MAIL = /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/;
  var ZW = /[​﻿]/g;

  // ------------------------------------------------------------ journal
  var J = (window.__journal = []);
  function noter(type, o) { var e = { n: J.length, t: Math.round(performance.now()), type: type }; for (var k in o) e[k] = o[k]; J.push(e); }
  function ident(el) {
    if (!el) return "?";
    if (el.nodeType === 9) return "document";
    if (el.nodeType !== 1) el = el.parentElement;
    if (!el) return "?";
    if (el.closest('[data-pp-panneau],[data-testid="ao-panel"]')) return "PANNEAU";
    var x = el.closest("[data-ident]");
    if (x) return x.getAttribute("data-ident");
    if (el === document.body) return "body";
    if (el === H) return "html";
    return el.tagName.toLowerCase() + (el.id ? "#" + el.id : "");
  }
  window.__ident = ident;
  window.__typed = []; window.__keys = [];
  document.addEventListener("click", function (e) { noter("clic", { cible: ident(e.target), confiance: e.isTrusted }); }, true);
  document.addEventListener("keydown", function (e) {
    var c = ident(e.target);
    if (c === "A") window.__keys.push(e.key);
    noter("touche", { cible: c, touche: e.key, confiance: e.isTrusted });
  }, true);
  document.addEventListener("input", function (e) { noter("saisie", { cible: ident(e.target), genre: e.inputType || "", data: e.data == null ? null : String(e.data).slice(0, 160), confiance: e.isTrusted }); }, true);
  document.addEventListener("paste", function (e) {
    var t = ""; try { t = e.clipboardData ? e.clipboardData.getData("text/plain") : ""; } catch (x) {}
    noter("collage", { cible: ident(e.target), texte: String(t).slice(0, 400), confiance: e.isTrusted });
  }, true);
  document.addEventListener("focusin", function (e) { noter("focus", { cible: ident(e.target) }); }, true);
  document.addEventListener("ao-run-begin", function () { window.__typed.length = 0; window.__keys.length = 0; seg = ""; noter("releve-debut", {}); });

  // ------------------------------------------------------------ annuaire
  var ANNUAIRE_DEFAUT = [
    ["MARTIN Alice", "alice.martin@ecole-test.fr"],
    ["DUPONT Jean", "jean.dupont@ecole-test.fr"],
    ["BERNARD Camille", "camille.bernard@ecole-test.fr"],
    ["PETIT Thomas", "thomas.petit@ecole-test.fr"],
    ["ROBERT Paul", "paul.robert@ecole-test.fr"],
    ["RICHARD Emma", "emma.richard@ecole-test.fr"],
    ["DURAND Léa", "lea.durand@ecole-test.fr"],
    ["MOREAU Hugo", "hugo.moreau@ecole-test.fr"],
    ["SIMON Chloé", "chloe.simon@ecole-test.fr"],
    ["LAURENT Louis", "louis.laurent@ecole-test.fr"],
    ["MICHEL Manon", "manon.michel@ecole-test.fr"],
    ["GARCIA Lucas", "lucas.garcia@ecole-test.fr"],
    ["Alice PRUNIER", "alice.prunier@partenaire.example", true],
    ["Arthur REMY", "arthur.remy@edu.example.com", true],
    ["Miriam ABDER", "abder@example.com", true],
    ["Scolarité Paris", "scolarite@ecole-test.fr"],
  ];
  var ANNUAIRE = (CFG.annuaire || ANNUAIRE_DEFAUT).map(function (c, i) { return { nom: c[0], adresse: c[1], externe: !!c[2], poids: c[3] != null ? c[3] : 1000 - i, i: i }; });
  function norm(s) { return String(s == null ? "" : s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
  function jetons(s) { return norm(s).split(/[^a-z0-9]+/).filter(Boolean); }
  function parPoids(a, b) { return b.poids - a.poids || a.i - b.i; }
  function chercher(q) {
    var qw = jetons(q);
    if (!qw.length) return [];
    var ql = String(q).trim().toLowerCase();
    return ANNUAIRE.filter(function (c) {
      var t = jetons(c.nom + " " + c.adresse);
      return qw.every(function (w) { return t.some(function (x) { return x.indexOf(w) === 0; }); });
    }).sort(function (a, b) {
      var ea = a.adresse.toLowerCase() === ql ? 1 : 0, eb = b.adresse.toLowerCase() === ql ? 1 : 0;
      return eb - ea || parPoids(a, b);
    });
  }
  function parAdresse(a) { var l = String(a).toLowerCase(); return ANNUAIRE.filter(function (c) { return c.adresse.toLowerCase() === l; })[0] || null; }
  function resoudre(t) {
    t = String(t).replace(ZW, "").trim();
    var m = t.match(/<([^>]+)>/);
    if (m) t = m[1].trim();
    var c = parAdresse(t);
    if (c) return { adresse: c.adresse, nom: c.nom, resolu: true };
    if (MAIL.test(t)) return { adresse: t, nom: null, resolu: true };
    return { adresse: null, nom: null, texte: t, resolu: false };
  }
  function estExterne(adresse) { return !/@ecole-test\.fr$/i.test(adresse); }

  // ------------------------------------------------------------ pastilles
  var INFOS = new WeakMap();
  var nPastille = 0xad;
  function creerPastille(c, source) {
    var s = document.createElement("span");
    s.id = "REK" + (++nPastille).toString(16).padStart(6, "0");
    s.className = "_Entity _EType_RECIPIENT_ENTITY _EReadonly_1 Lbs4W";
    s.setAttribute("contenteditable", "false");
    s.setAttribute("draggable", "true");
    var aff = c.adresse || c.texte;
    s.setAttribute("aria-label", "offline" + aff);
    s.innerHTML = '<span class="lpcWrapper ___1kqcvn5"><span class="pill-222 ' + (c.resolu ? "validPill-223" : "invalidPill-224") +
      '"><span class="presence"><span id="presenceIcon" role="img" aria-label="offline"></span></span><span class="pillText"></span></span></span>';
    s.querySelector(".pillText").textContent = aff;
    INFOS.set(s, { adresse: c.adresse || null, nom: c.nom || null, texte: c.texte || null, resolu: !!c.resolu, source: source });
    if (c.nom) {
      var poser = function () { if (!s.isConnected) return; s.setAttribute("aria-label", "offline" + c.nom); s.querySelector(".pillText").textContent = c.nom; };
      if (CFG.delaiResolution <= 0) poser(); else setTimeout(poser, D(CFG.delaiResolution));
    }
    return s;
  }
  function noeudsLibres(champ) {
    var out = [], w = document.createTreeWalker(champ, NodeFilter.SHOW_TEXT);
    for (var n = w.nextNode(); n; n = w.nextNode()) {
      if (n.parentElement && n.parentElement.closest("._EType_RECIPIENT_ENTITY")) continue;
      out.push(n);
    }
    return out;
  }
  function texteLibre(champ) { return noeudsLibres(champ).map(function (n) { return n.nodeValue; }).join("").replace(ZW, ""); }
  function retirerTexteLibre(champ) {
    noeudsLibres(champ).forEach(function (n) { n.remove(); });
    [].slice.call(champ.querySelectorAll("br")).forEach(function (b) { if (!b.closest("._EType_RECIPIENT_ENTITY")) b.remove(); });
  }
  function pastillesDe(champ) { return [].slice.call(champ.querySelectorAll("._EType_RECIPIENT_ENTITY")); }
  function placerFin(champ) {
    if (document.activeElement !== champ) return;
    var r = document.createRange(); r.selectNodeContents(champ); r.collapse(false);
    var s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }
  function ajouterPastilles(champ, contacts, source) {
    contacts.forEach(function (c) {
      champ.appendChild(document.createTextNode("​"));
      champ.appendChild(creerPastille(c, source));
      noter("pastille", { champ: champ.getAttribute("data-ident"), adresse: c.adresse || null, texte: c.texte || null, nom: c.nom || null, source: source });
    });
    champ.appendChild(document.createTextNode("​"));
    placerFin(champ);
  }
  function infosChamp(champ) {
    if (!champ) return [];
    return pastillesDe(champ).map(function (p) { var i = INFOS.get(p) || {}; return { adresse: i.adresse, nom: i.nom, texte: i.texte, resolu: i.resolu, source: i.source, affiche: (p.textContent || "").trim() }; });
  }

  // ------------------------------------------------------------ suggestions du champ « À »
  var POP = { el: null, champ: null, q: "", liste: [], montres: 0, surligne: 0 };
  var tSugg = null, seg = "";
  function popup() {
    if (POP.el) return POP.el;
    var d = document.createElement("div");
    d.className = "fui-PopoverSurface"; d.id = "suggestions-destinataires"; d.setAttribute("role", "presentation"); d.hidden = true;
    d.innerHTML = '<div role="listbox" aria-label="Contacts suggérés" id="liste-suggestions"></div><div class="pied-zone"></div>';
    // Comme OWA : cliquer dans la liste ne fait pas perdre le focus au champ.
    d.addEventListener("mousedown", function (e) { e.preventDefault(); });
    document.getElementById("calques").appendChild(d);
    POP.el = d;
    return d;
  }
  function initiales(nom) { var p = String(nom).split(/\s+/).filter(Boolean); return (((p[0] || "")[0] || "") + ((p[1] || "")[0] || "")).toUpperCase(); }
  function optionContact(c, i, sel) {
    var o = document.createElement("div");
    o.setAttribute("role", "option"); o.id = "sugg-opt-" + i; o.className = "fui-Option";
    o.setAttribute("data-ident", "suggestion:" + c.adresse);
    o.setAttribute("aria-selected", String(!!sel));
    o.setAttribute("aria-label", c.nom + ", " + c.adresse + (c.externe ? ", Externe" : ""));
    o.innerHTML = '<div class="fui-Persona"><span class="fui-Avatar"><span class="fui-Avatar__initials"></span></span><span class="fui-Persona__primaryText"></span>' +
      (c.externe ? '<span class="badge">Externe</span>' : "") + '<span class="fui-Persona__secondaryText"></span></div>';
    o.querySelector(".fui-Avatar__initials").textContent = initiales(c.nom);
    o.querySelector(".fui-Persona__primaryText").textContent = c.nom;
    var sec = o.querySelector(".fui-Persona__secondaryText"); sec.textContent = c.adresse; sec.title = c.adresse;
    o.addEventListener("click", function () { choisirSuggestion(c); });
    return o;
  }
  function rendreSuggestions() {
    var d = popup(), lb = d.querySelector('[role="listbox"]'), pied = d.querySelector(".pied-zone");
    var vis = POP.liste.slice(0, POP.montres);
    lb.innerHTML = "";
    vis.forEach(function (c, i) { lb.appendChild(optionContact(c, i, i === POP.surligne)); });
    pied.innerHTML = "";
    if (POP.q && POP.liste.length > POP.montres) {
      var m = document.createElement("div");
      m.setAttribute("role", "button"); m.tabIndex = 0; m.className = "pied-sugg"; m.setAttribute("data-ident", "plus-de-resultats");
      m.innerHTML = '<i class="icone" aria-hidden="true"></i><span>Afficher plus de résultats</span>';
      m.addEventListener("click", function () { POP.montres += CFG.lot; noter("plus-de-resultats", { q: POP.q, montres: POP.montres }); rendreSuggestions(); });
      pied.appendChild(m);
    }
    var vide = !vis.length && !pied.childElementCount;
    d.hidden = vide;
    if (!vide && POP.champ) {
      var r = POP.champ.getBoundingClientRect();
      d.style.left = Math.round(r.left) + "px"; d.style.top = Math.round(r.bottom + 2) + "px";
    }
  }
  function fermerSuggestions() { clearTimeout(tSugg); if (POP.el) POP.el.hidden = true; POP.champ = null; POP.liste = []; POP.q = ""; }
  function majSuggestions(champ) {
    if (!champ.isConnected) { fermerSuggestions(); return; }
    var q = texteLibre(champ).trim();
    if (!q) {
      if (CFG.suggestionsAuFocus && document.activeElement === champ) {
        POP.champ = champ; POP.q = ""; POP.liste = ANNUAIRE.slice().sort(parPoids).slice(0, CFG.initialFocus); POP.montres = POP.liste.length; POP.surligne = 0;
        rendreSuggestions();
      } else fermerSuggestions();
      return;
    }
    var nouvelle = q !== POP.q || POP.champ !== champ;
    POP.champ = champ; POP.q = q; POP.liste = chercher(q);
    if (nouvelle) { POP.montres = CFG.initial; POP.surligne = 0; }
    rendreSuggestions();
  }
  function planifierSuggestions(champ) { clearTimeout(tSugg); tSugg = setTimeout(function () { majSuggestions(champ); }, D(CFG.delaiSuggestions)); }
  function choisirSuggestion(c) {
    var champ = POP.champ;
    if (!champ) return;
    retirerTexteLibre(champ);
    ajouterPastilles(champ, [{ adresse: c.adresse, nom: c.nom, resolu: true }], "clic sur une suggestion");
    fermerSuggestions();
  }
  function suggestionSurlignee(champ) {
    if (POP.champ !== champ || !POP.el || POP.el.hidden || !POP.q) return null;
    return POP.liste.slice(0, POP.montres)[POP.surligne] || null;
  }
  function suivreFrappe(champ) {
    if (champ.getAttribute("data-ident") !== "A") return;
    var v = texteLibre(champ).trim();
    if (v === "") { if (seg) { window.__typed.push(seg); seg = ""; } }
    else if (v.length >= seg.length) seg = v;
  }

  // Comportement d'un champ de destinataires (À, Cc, Cci).
  function brancherPuits(champ) {
    champ.addEventListener("focus", function () { planifierSuggestions(champ); });
    champ.addEventListener("keydown", function (e) {
      var k = e.key;
      if ((k === "ArrowDown" || k === "ArrowUp") && suggestionSurlignee(champ)) {
        e.preventDefault();
        POP.surligne = Math.max(0, Math.min(POP.montres - 1, POP.surligne + (k === "ArrowDown" ? 1 : -1)));
        rendreSuggestions();
        return;
      }
      if (k === ";" || k === "," || k === "Enter" || k === "Tab") {
        var libre = texteLibre(champ).trim();
        if (!libre) return;
        e.preventDefault();
        var s = suggestionSurlignee(champ);
        retirerTexteLibre(champ);
        ajouterPastilles(champ, [s ? { adresse: s.adresse, nom: s.nom, resolu: true } : resoudre(libre)], "touche " + k);
        fermerSuggestions();
      }
    });
    champ.addEventListener("input", function () {
      suivreFrappe(champ);
      var libre = texteLibre(champ);
      if (/[;,]/.test(libre)) {
        var segs = libre.split(/[;,]/), reste = segs.pop();
        var contacts = segs.map(function (s) { return s.trim(); }).filter(Boolean).map(function (s) {
          if (CFG.pointVirgule === "suggestion") { var b = chercher(s)[0]; if (b) return { adresse: b.adresse, nom: b.nom, resolu: true }; }
          return resoudre(s);
        });
        retirerTexteLibre(champ);
        ajouterPastilles(champ, contacts, "saisie avec « ; »");
        if (reste.trim()) { champ.appendChild(document.createTextNode(reste.trim())); placerFin(champ); }
        suivreFrappe(champ);
        fermerSuggestions();
        return;
      }
      planifierSuggestions(champ);
    });
    champ.addEventListener("paste", function (e) {
      if ((CFG.collageIgnore || []).indexOf(champ.getAttribute("data-ident")) >= 0) return;
      var t = e.clipboardData ? e.clipboardData.getData("text/plain") : "";
      if (!t) return;
      e.preventDefault();
      var tout = texteLibre(champ) + t;
      var finit = /[;,\n]\s*$/.test(t);
      var parts = tout.split(/[;,\n]+/).map(function (s) { return s.trim(); }).filter(Boolean);
      var reste = "";
      if (!finit && parts.length && !MAIL.test(parts[parts.length - 1].replace(/^.*<|>$/g, ""))) reste = parts.pop();
      retirerTexteLibre(champ);
      ajouterPastilles(champ, parts.map(resoudre), "collage");
      if (reste) { champ.appendChild(document.createTextNode(reste)); placerFin(champ); }
      planifierSuggestions(champ);
    });
    // Outlook résout ce qui reste tapé quand on quitte le champ (variante réglable).
    champ.addEventListener("focusout", function (e) {
      var vers = e.relatedTarget;
      if (vers && POP.el && POP.el.contains(vers)) return;
      setTimeout(function () {
        if (document.activeElement === champ || !champ.isConnected) return;
        var libre = texteLibre(champ).trim();
        if (libre && CFG.flou === "pastille") {
          retirerTexteLibre(champ);
          ajouterPastilles(champ, [resoudre(libre)], "perte du focus");
          suivreFrappe(champ);
        }
        if (POP.champ === champ) fermerSuggestions();
      }, D(40));
    });
  }

  // ------------------------------------------------------------ menus et fenêtres
  var MENUS = [];
  function fermerMenus() { MENUS.splice(0).forEach(function (m) { m.remove(); }); [].forEach.call(document.querySelectorAll('[aria-expanded="true"][aria-haspopup]'), function (b) { b.setAttribute("aria-expanded", "false"); }); }
  function ouvrirMenu(ancre, items, nom) {
    fermerMenus();
    var m = document.createElement("div");
    m.setAttribute("role", "menu"); m.className = "fui-MenuPopover"; m.setAttribute("data-ident", "menu:" + nom);
    items.forEach(function (it) {
      var b = document.createElement("button");
      b.type = "button"; b.setAttribute("role", it.role || "menuitem"); b.setAttribute("data-ident", it.ident);
      if (it.coche != null) b.setAttribute("aria-checked", String(!!it.coche()));
      b.textContent = it.texte;
      b.addEventListener("click", function () { noter("menu", { item: it.ident }); fermerMenus(); if (it.action) it.action(); });
      m.appendChild(b);
    });
    var r = ancre.getBoundingClientRect();
    m.style.left = Math.max(0, Math.min(innerWidth - 240, Math.round(r.left))) + "px"; m.style.top = Math.round(r.bottom + 2) + "px";
    document.getElementById("calques").appendChild(m);
    MENUS.push(m);
    if (ancre.hasAttribute("aria-haspopup")) ancre.setAttribute("aria-expanded", "true");
    noter("menu-ouvert", { nom: nom });
    return m;
  }
  document.addEventListener("keydown", function (e) { if (e.key === "Escape" && MENUS.length) fermerMenus(); }, true);
  document.addEventListener("mousedown", function (e) { if (!e.isTrusted) return; if (MENUS.length && !MENUS.some(function (m) { return m.contains(e.target); }) && !(e.target.closest && e.target.closest("[aria-haspopup]"))) fermerMenus(); }, true);

  var MODALES = [], nDlg = 10;
  function ouvrirModale(o) {
    var n = ++nDlg;
    var calque = document.createElement("div");
    calque.className = "ms-Layer ms-Layer--fixed";
    calque.innerHTML = '<div role="' + (o.alerte ? "alertdialog" : "dialog") + '" aria-modal="true" aria-labelledby="Dialog' + n + '-title" data-ident="dialogue:' + (o.id || n) + '">' +
      '<div class="ms-Modal is-open ms-Dialog root-' + n + '" tabindex="-1"><div class="ms-Overlay ms-Overlay--dark"></div><div class="ms-Dialog-main">' +
      '<div class="ms-Dialog-title" id="Dialog' + n + '-title"></div><div class="ms-Dialog-inner"><div class="texte"></div><div class="contenu"></div></div><div class="ms-Dialog-actions"></div></div></div></div>';
    calque.querySelector(".ms-Dialog-title").textContent = o.titre || "";
    calque.querySelector(".texte").textContent = o.texte || "";
    if (o.contenu) calque.querySelector(".contenu").appendChild(o.contenu);
    if (o.champ) { var ta = document.createElement("textarea"); ta.setAttribute("aria-label", o.champ); ta.setAttribute("data-ident", "dialogue:" + (o.id || n) + ":champ"); calque.querySelector(".contenu").appendChild(ta); }
    var m = { calque: calque, piege: o.piege !== false, avant: document.activeElement, o: o };
    (o.boutons || []).forEach(function (bt) {
      var b = document.createElement("button");
      b.type = "button"; b.textContent = bt.texte; b.setAttribute("data-ident", bt.ident);
      b.addEventListener("click", function () { noter("dialogue-bouton", { bouton: bt.ident, confiance: true }); if (bt.fermer !== false) fermerModale(m); if (bt.action) bt.action(); });
      calque.querySelector(".ms-Dialog-actions").appendChild(b);
    });
    document.getElementById("calques").appendChild(calque);
    MODALES.push(m);
    var premier = calque.querySelector("textarea,input,button") || calque.querySelector(".ms-Modal");
    try { premier.focus(); } catch (e) {}
    noter("modale", { id: o.id || n, ouverte: true });
    return m;
  }
  function fermerModale(m) {
    if (!m || !m.calque.isConnected) return;
    m.calque.remove();
    MODALES.splice(MODALES.indexOf(m), 1);
    noter("modale", { id: m.o.id, ouverte: false });
    // Fluent rend le focus à l'élément qui l'avait avant la fenêtre.
    if (!MODALES.length && m.avant && m.avant.isConnected && m.avant !== document.body) { try { m.avant.focus(); } catch (e) {} }
  }
  // FocusTrapZone de Fluent (Modal) : tant que la fenêtre est là, le focus y revient.
  document.addEventListener("focusin", function (e) {
    var m = MODALES[MODALES.length - 1];
    if (!m || !m.piege || m.calque.contains(e.target)) return;
    var b = m.calque.querySelector("textarea,input,button") || m.calque.querySelector(".ms-Modal");
    if (b && b !== e.target) { noter("piege-focus", { voulu: ident(e.target) }); b.focus(); }
  }, true);

  function ouvrirCarnet(c, ligne) {
    window.__carnet = (window.__carnet || 0) + 1;
    H.setAttribute("data-carnet", String(window.__carnet));
    noter("carnet", { ligne: ligne });
    var contenu = document.createElement("div");
    contenu.innerHTML = '<input type="text" aria-label="Rechercher des personnes" placeholder="Rechercher" data-ident="carnet:recherche">' +
      '<div role="listbox" aria-label="Contacts" data-ident="carnet:liste">' + ANNUAIRE.map(function (x) { return '<div role="option" data-ident="carnet:contact"><span>' + echapper(x.nom) + '</span> <span>' + echapper(x.adresse) + "</span></div>"; }).join("") + "</div>" +
      (CFG.carnetChamps ? '<div>À <div contenteditable="true" aria-label="À" data-ident="carnet:champ-a" style="min-height:20px;border:1px solid #ccc"></div></div><div>Cci <div contenteditable="true" aria-label="Cci" data-ident="carnet:champ-cci" style="min-height:20px;border:1px solid #ccc"></div></div>' : "");
    ouvrirModale({ id: "carnet", titre: "Ajouter des destinataires", contenu: contenu, boutons: [
      { texte: "À ->", ident: "carnet:a", fermer: false }, { texte: "Cc ->", ident: "carnet:cc", fermer: false }, { texte: "Cci ->", ident: "carnet:cci", fermer: false },
      { texte: "Enregistrer", ident: "carnet:enregistrer" }, { texte: "Annuler", ident: "carnet:annuler" },
    ] });
  }

  // ------------------------------------------------------------ données de la boîte
  var MESSAGES = CFG.messages || [
    { de: "Apple", adresse: "noreply@email.apple.com", objet: "Votre reçu Apple", apercu: "Vous n'obtenez pas souvent d'e-mail à partir de noreply@email.apple.com", corps: "Merci pour votre achat. Questions : support@apple.example" },
    { de: "MARTIN Alice", adresse: "alice.martin@ecole-test.fr", objet: "Réunion de rentrée", apercu: "Bonjour, la réunion aura lieu lundi. Contact : secretariat@ecole-test.fr", corps: "Bonjour,\nLa réunion de rentrée aura lieu lundi à 9 h.\nPour toute question : secretariat@ecole-test.fr ou direction.campus@ecole-test.fr\nAlice" },
    { de: "Microsoft Learn", adresse: "learn@mails.microsoft.com", objet: "Nouveautés de la semaine", apercu: "learn@mails.microsoft.com vous présente les nouveautés", corps: "Découvrez les nouveautés." },
  ];
  var envoyes = (window.envoyes = []);
  var BOITE = (window.__boiteEnvoi = []);
  function publier() { H.setAttribute("data-envoyes", JSON.stringify(envoyes)); H.setAttribute("data-boite-envoi", String(BOITE.length)); }
  publier();

  // ------------------------------------------------------------ ruban
  var ONGLETS_BASE = ["Accueil", "Afficher", "Aide"];
  var ONGLETS_REDACTION = ["Message", "Insérer", "Mettre le texte en forme", "Dessiner", "Options"];
  var ongletRuban = "Accueil";
  function bouton(identifiant, libelle, opts) {
    opts = opts || {};
    var b = document.createElement("button");
    b.type = "button"; b.setAttribute("data-ident", identifiant);
    if (opts.aria !== false) b.setAttribute("aria-label", opts.aria || libelle);
    if (opts.id) b.id = opts.id;
    if (opts.pressed != null) b.setAttribute("aria-pressed", String(opts.pressed));
    if (opts.popup) { b.setAttribute("aria-haspopup", "menu"); b.setAttribute("aria-expanded", "false"); }
    b.innerHTML = '<i class="icone" aria-hidden="true"></i>';
    if (opts.icone) b.appendChild(document.createTextNode(opts.icone)); else b.appendChild(document.createTextNode(libelle));
    if (opts.action) b.addEventListener("click", opts.action);
    return b;
  }
  function boutonImportance(niveau) {
    var c = ACTIVE;
    var haute = niveau === "haute";
    var lib = haute ? "Importance élevée" : "Importance faible";
    return bouton(haute ? "importance-haute" : "importance-faible", lib, { pressed: CFG.importanceSansEtat ? null : !!(c && c.importanceAffichee === niveau), action: function () { basculerImportance(niveau); } });
  }
  function basculerImportance(niveau) {
    var c = ACTIVE;
    if (!c) return;
    c.importance = c.importance === niveau ? "normale" : niveau;
    noter("importance", { valeur: c.importance });
    var v = c.importance;
    setTimeout(function () { c.importanceAffichee = v; if (ACTIVE === c) majImportance(); }, D(CFG.importanceDelai));
  }
  function majImportance() {
    var c = ACTIVE;
    [["importance-haute", "haute"], ["importance-faible", "faible"]].forEach(function (x) {
      var b = document.querySelector('#commandes [data-ident="' + x[0] + '"]');
      if (b && !CFG.importanceSansEtat) b.setAttribute("aria-pressed", String(!!(c && c.importanceAffichee === x[1])));
    });
  }
  function itemsDebordement() {
    var items = [];
    if (CFG.rubanEtroit && ongletRuban === "Message") {
      items.push({ texte: "Importance élevée", ident: "debord:importance-haute", role: "menuitemcheckbox", coche: function () { return ACTIVE && ACTIVE.importanceAffichee === "haute"; }, action: function () { basculerImportance("haute"); } });
      items.push({ texte: "Importance faible", ident: "debord:importance-faible", role: "menuitemcheckbox", coche: function () { return ACTIVE && ACTIVE.importanceAffichee === "faible"; }, action: function () { basculerImportance("faible"); } });
      items.push({ texte: "Sensibilité", ident: "debord:sensibilite" }, { texte: "Chiffrer", ident: "debord:chiffrer" }, { texte: "Imprimer un brouillon", ident: "debord:imprimer" });
    }
    items.push({ texte: "Dicter", ident: "debord:dicter" }, { texte: "Éditeur", ident: "debord:editeur" }, { texte: "Afficher De", ident: "debord:afficher-de" }, { texte: "Options du message", ident: "debord:options-message" });
    return items;
  }
  function rendreRuban() {
    var redaction = !!composeVisible();
    var onglets = redaction ? ONGLETS_BASE.concat(ONGLETS_REDACTION) : ONGLETS_BASE;
    if (onglets.indexOf(ongletRuban) < 0) ongletRuban = "Accueil";
    var tl = document.getElementById("onglets-ruban");
    tl.innerHTML = "";
    onglets.forEach(function (t) {
      var b = document.createElement("button");
      b.type = "button"; b.setAttribute("role", "tab"); b.setAttribute("aria-selected", String(t === ongletRuban)); b.setAttribute("data-ident", "onglet:" + t);
      b.textContent = t;
      b.addEventListener("click", function () { ongletRuban = t; rendreRuban(); });
      tl.appendChild(b);
    });
    var cm = document.getElementById("commandes");
    cm.innerHTML = "";
    var add = function (el) { cm.appendChild(el); };
    if (ongletRuban === "Accueil") {
      var sp = document.createElement("span"); sp.className = "fui-SplitButton";
      var nb = document.createElement("button"); nb.type = "button"; nb.className = "fui-Button fui-SplitButton__primaryActionButton"; nb.setAttribute("data-ident", "nouveau");
      nb.innerHTML = '<i class="icone" aria-hidden="true"></i>'; nb.appendChild(document.createTextNode("Nouveau message"));
      nb.addEventListener("click", function () { nouvelleRedaction({ type: "nouveau", focus: "A" }); });
      sp.appendChild(nb);
      sp.appendChild(bouton("nouveau-options", "⌄", { aria: "Autres options de nouveau message", popup: true, action: function () { ouvrirMenu(this, [{ texte: "Courrier", ident: "nouveau:courrier" }, { texte: "Événement", ident: "nouveau:evenement" }], "nouveau"); } }));
      add(sp);
      [["ruban:supprimer", "Supprimer"], ["ruban:archiver", "Archiver"], ["ruban:signaler", "Signaler"], ["ruban:deplacer", "Déplacer vers"], ["ruban:repondre", "Répondre"],
        ["ruban:repondre-tous", "Répondre à tous"], ["ruban:transferer", "Transférer"], ["ruban:non-lu", "Marquer comme non lu"], ["ruban:indicateur", "Indicateur"]]
        .forEach(function (x) { add(bouton(x[0], x[1], { action: x[0] === "ruban:repondre" ? function () { if (CFG.lecture != null) repondre(lectureCourante); } : null })); });
    } else if (ongletRuban === "Message") {
      add(bouton("ruban:carnet", "Carnet d'adresses", { action: function () { ouvrirCarnet(ACTIVE, "ruban"); } }));
      add(bouton("ruban:verifier-noms", "Vérifier les noms"));
      add(bouton("ruban:joindre", "Joindre un fichier"));
      add(bouton("ruban:lien", "Lien"));
      add(bouton("ruban:signature", "Signature", { popup: true, action: function () { ouvrirMenu(this, [{ texte: "Signatures…", ident: "signature:gerer" }], "signature"); } }));
      add(bouton("ruban:enregistrer", "Enregistrer"));
      if (!CFG.rubanEtroit) {
        add(boutonImportance("haute"));
        add(boutonImportance("faible"));
        add(bouton("ruban:sensibilite", "Sensibilité"));
        add(bouton("ruban:chiffrer", "Chiffrer"));
        add(bouton("ruban:imprimer", "Imprimer un brouillon"));
      }
    } else if (ongletRuban === "Options") {
      [["options:themes", "Thèmes"], ["options:format", "Format"], ["options:confirmation", "Confirmation de lecture"], ["options:differer", "Différer la remise"]].forEach(function (x) { add(bouton(x[0], x[1])); });
    } else if (ongletRuban === "Insérer") {
      [["inserer:tableau", "Tableau"], ["inserer:images", "Images"], ["inserer:emoji", "Emoji"]].forEach(function (x) { add(bouton(x[0], x[1])); });
    } else if (ongletRuban === "Mettre le texte en forme") {
      [["forme:gras", "Gras"], ["forme:italique", "Italique"]].forEach(function (x) { add(bouton(x[0], x[1])); });
    } else if (ongletRuban === "Dessiner") {
      add(bouton("dessiner:stylet", "Stylet"));
    } else if (ongletRuban === "Afficher") {
      [["afficher:parametres", "Paramètres d'affichage"], ["afficher:disposition", "Disposition"]].forEach(function (x) { add(bouton(x[0], x[1])); });
    } else if (ongletRuban === "Aide") {
      [["aide:aide", "Aide"], ["aide:commentaires", "Commentaires"]].forEach(function (x) { add(bouton(x[0], x[1])); });
    }
    add(bouton("ruban:autres-options", "Autres options", { id: "RibbonOverflowMenu-overflow", popup: true, action: function () { ouvrirMenu(this, itemsDebordement(), "debordement-" + ongletRuban); } }));
  }

  // ------------------------------------------------------------ liste et volet de lecture
  var lectureCourante = CFG.lecture;
  function rendreListe() {
    var l = document.getElementById("liste");
    l.innerHTML = "";
    MESSAGES.forEach(function (m, i) {
      var d = document.createElement("div");
      d.className = "msg"; d.setAttribute("role", "option"); d.setAttribute("aria-selected", String(i === lectureCourante)); d.setAttribute("data-ident", "liste:" + i);
      d.setAttribute("aria-label", m.de + " " + m.objet + " " + m.apercu);
      d.innerHTML = '<div class="de"></div><div class="objet"></div><div class="apercu"></div><div class="actions"></div>';
      d.querySelector(".de").textContent = m.de; d.querySelector(".objet").textContent = m.objet; d.querySelector(".apercu").textContent = m.apercu;
      // Actions au survol : icônes seules (libellé dans aria-label), comme le vrai Outlook.
      [["supprimer", "Supprimer", "🗑"], ["non-lu", "Marquer comme non lu", "✉"], ["indicateur", "Indicateur", "⚑"], ["epingler", "Épingler", "📌"]].forEach(function (x) {
        d.querySelector(".actions").appendChild(bouton("liste:" + i + ":" + x[0], x[1], { icone: x[2] }));
      });
      d.addEventListener("click", function (e) { if (e.target.closest(".actions")) return; selectionner(i); });
      l.appendChild(d);
    });
  }
  function rendreLecture() {
    var z = document.getElementById("lecture");
    z.innerHTML = "";
    var m = lectureCourante != null ? MESSAGES[lectureCourante] : null;
    if (!m) { z.innerHTML = '<div class="vide">Sélectionnez un élément à lire</div>'; return; }
    var d = document.createElement("div");
    d.setAttribute("data-ident", "lecture");
    d.innerHTML = '<h2 class="objet-lu"></h2><div class="entete-lu"><span class="fui-Persona__primaryText nom"></span> <span class="adresse-lu"></span><div>À : Moi &lt;' + echapper(CFG.moi) + '&gt;</div></div><div class="actions-lu"></div><div class="corps-lu" style="white-space:pre-wrap"></div>';
    d.querySelector(".objet-lu").textContent = m.objet;
    d.querySelector(".nom").textContent = m.de;
    var a = d.querySelector(".adresse-lu"); a.textContent = "<" + m.adresse + ">"; a.title = m.adresse;
    d.querySelector(".corps-lu").textContent = m.corps;
    var act = d.querySelector(".actions-lu");
    act.appendChild(bouton("lecture:repondre", "Répondre", { action: function () { repondre(lectureCourante); } }));
    act.appendChild(bouton("lecture:repondre-tous", "Répondre à tous"));
    act.appendChild(bouton("lecture:transferer", "Transférer"));
    z.appendChild(d);
  }
  function selectionner(i) {
    lectureCourante = i;
    if (ACTIVE) { ACTIVE.visible = false; }
    afficherZone();
    rendreListe();
    rendreLecture();
  }

  // ------------------------------------------------------------ rédactions
  var COMPOSES = [], ACTIVE = null, nMsg = 0, nIds = 11, nSplit = 200;
  function composeVisible() { return ACTIVE && ACTIVE.visible && ACTIVE.el.isConnected ? ACTIVE : null; }
  window.__composeVisible = function () { return !!composeVisible(); };
  function ligneHTML(mid, cle, lib, labelId, msgSuffix, avecBascules) {
    return '<div class="ligne-dest" data-ligne="' + cle + '">' +
      '<div class="libelle"><button type="button" data-is-focusable="true" tabindex="0" role="button" aria-roledescription="bouton, ouvre la boîte de dialogue Ajouter des destinataires" class="fui-Button r1f29ykk HnZn5" id="recipient-well-label-' + labelId + '" data-ident="libelle-' + cle.toLowerCase() + '">' + lib + "</button></div>" +
      '<span class="fui-Input r1jtohuq ___1v6zkb4"><div class="LCArn ___1vrlxok" id="MSG_' + mid + "_" + msgSuffix + '" role="group">' +
      '<div aria-label="' + lib + '" tabindex="0" class="___h7sz9k0 f22iagw f1vx9l62 EditorClass" spellcheck="false" contenteditable="true" id="' + (++nIds) + '" data-ident="' + cle + '"></div></div>' +
      '<input type="text" class="fui-Input__input r1j2f4yb" value=""></span>' +
      (avecBascules ? '<span class="bascules"><button type="button" class="fui-Button" data-ident="bascule-cc" title="Afficher Cc">Cc</button><button type="button" class="fui-Button" data-ident="bascule-cci" title="Afficher Cci">Cci</button></span>' : "") +
      "</div>";
  }
  function nouvelleRedaction(o) {
    o = o || {};
    if (ACTIVE) ACTIVE.visible = false;
    var mid = "AAQkADI" + (++nMsg) + "Zj";
    var c = { id: mid, type: o.type || "nouveau", importance: "normale", importanceAffichee: "normale", visible: true, ouverte: performance.now() };
    var el = document.createElement("div");
    el.className = "redaction"; el.setAttribute("data-compose", mid);
    var sid = ++nSplit;
    var objetVisible = o.objetVisible !== false;
    el.innerHTML =
      '<div class="barre" role="toolbar" aria-label="Rédaction">' +
      '<span class="fui-SplitButton"><button type="button" id="splitButton-r' + sid + '__primaryActionButton" aria-label="Envoyer" class="fui-Button r1alrhcs fui-SplitButton__primaryActionButton" data-ident="envoyer">Envoyer</button>' +
      '<button type="button" id="splitButton-r' + sid + '__menuButton" aria-label="Autres options d\'envoi" aria-haspopup="menu" aria-expanded="false" class="fui-Button fui-MenuButton fui-SplitButton__menuButton" data-ident="envoyer-options">⌄</button></span>' +
      '<button type="button" aria-label="Réglez votre sensibilité" aria-haspopup="menu" data-ident="sensibilite">Général</button>' +
      '<span style="flex:1"></span>' +
      '<button type="button" aria-label="Abandonner" data-ident="abandonner">🗑</button>' +
      '<button type="button" aria-label="Isoler" data-ident="isoler">⧉</button></div>' +
      '<div class="ligne-de">De : ' + echapper(CFG.moi) + "</div>" +
      ligneHTML(mid, "A", "À", "to", "TO", true) +
      ligneHTML(mid, "CC", "Cc", "cc", "CC", false) +
      ligneHTML(mid, "CCI", "Cci", "bcc", "BCC", false) +
      '<div class="ligne-objet"' + (objetVisible ? "" : " hidden") + '><input type="text" maxlength="255" autocomplete="off" spellcheck="true" placeholder="Ajouter un objet" aria-label="Objet" id="MSG_' + mid + '_SUBJECT" class="fui-Input__input r12stul0" data-ident="OBJET"></div>' +
      '<div class="bandeau-redaction"></div>' +
      '<div tabindex="0" dir="ltr" class="dFCbN customScrollBar" role="textbox" aria-multiline="true" aria-label="Corps du message" contenteditable="true" data-ms-editor="true" data-scrollable="true" data-ident="CORPS">' +
      '<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div></div>';
    c.el = el;
    c.puits = { A: el.querySelector('[data-ident="A"]'), CC: el.querySelector('[data-ident="CC"]'), CCI: el.querySelector('[data-ident="CCI"]') };
    c.lignes = { CC: el.querySelector('[data-ligne="CC"]'), CCI: el.querySelector('[data-ligne="CCI"]') };
    c.objet = el.querySelector('[data-ident="OBJET"]');
    c.corps = el.querySelector('[data-ident="CORPS"]');
    c.lignes.CC.hidden = true;
    c.lignes.CCI.hidden = !CFG.cciVisible;
    if (CFG.cciVisible) el.querySelector('[data-ident="bascule-cci"]').remove();
    Object.keys(c.puits).forEach(function (k) { brancherPuits(c.puits[k]); });
    // Libellés « À », « Cc », « Cci » : ouvrent le carnet d'adresses (le piège).
    ["a", "cc", "cci"].forEach(function (k) { el.querySelector('[data-ident="libelle-' + k + '"]').addEventListener("click", function () { ouvrirCarnet(c, k); }); });
    el.querySelector('[data-ident="bascule-cc"]').addEventListener("click", function () { montrerLigne(c, "CC", this); });
    var bcci = el.querySelector('[data-ident="bascule-cci"]');
    if (bcci) bcci.addEventListener("click", function () { montrerLigne(c, "CCI", this); });
    el.querySelector('[data-ident="envoyer"]').addEventListener("click", function () { cliquerEnvoyer(c); });
    el.querySelector('[data-ident="envoyer-options"]').addEventListener("click", function () { noter("autres-options-envoi", {}); ouvrirMenu(this, [{ texte: "Planifier l'envoi", ident: "envoi:planifier" }], "envoi"); });
    el.querySelector('[data-ident="sensibilite"]').addEventListener("click", function () { ouvrirMenu(this, [{ texte: "Public", ident: "sensibilite:public" }, { texte: "Général", ident: "sensibilite:general" }, { texte: "Confidentiel", ident: "sensibilite:confidentiel" }], "sensibilite"); });
    el.querySelector('[data-ident="abandonner"]').addEventListener("click", function () { noter("abandon", {}); fermerRedaction(c, "abandonnee"); });
    el.querySelector('[data-ident="isoler"]').addEventListener("click", function () { noter("isoler", {}); fermerRedaction(c, "isolee"); });
    c.objet.addEventListener("input", majOngletsDocs);
    // Variante : une fenêtre d'Outlook avec un champ de saisie (enquête « Votre avis ») surgit
    // au moment où l'on entre dans l'objet ; piège à focus Fluent ; se referme seule.
    if (CFG.modaleSurObjet) {
      var vue = false;
      c.objet.addEventListener("focus", function () {
        if (vue) return; vue = true;
        var mo = ouvrirModale({ id: "enquete", titre: "Votre avis sur Outlook", texte: "Quelle est la probabilité que vous recommandiez Outlook ?", champ: "Commentaires", piege: true,
          boutons: [{ texte: "Envoyer", ident: "dialogue:enquete:envoyer" }, { texte: "Plus tard", ident: "dialogue:enquete:plus-tard" }] });
        setTimeout(function () { fermerModale(mo); }, D(CFG.modaleSurObjet.duree || 2000));
      });
    }
    // Éditeur du corps (Rooster) : il prend le collage en charge.
    c.corps.addEventListener("paste", function (e) {
      if ((CFG.collageIgnore || []).indexOf("CORPS") >= 0) return;
      var h = e.clipboardData && e.clipboardData.getData("text/html");
      var t = e.clipboardData && e.clipboardData.getData("text/plain");
      if (!h && !t) return;
      e.preventDefault();
      document.execCommand("insertHTML", false, h || echapper(t).replace(/\n/g, "<br>"));
      [].forEach.call(c.corps.querySelectorAll("a[href]"), function (a, i) { a.setAttribute("data-auth", "NotApplicable"); a.setAttribute("data-linkindex", String(i)); });
    });
    // Pré-remplissage (réponse, brouillon de l'utilisateur).
    if (o.a && o.a.length) ajouterPastilles(c.puits.A, o.a.map(resoudre), "pré-rempli");
    if (o.objet) c.objet.value = o.objet;
    if (o.corpsHtml) c.corps.innerHTML = o.corpsHtml;
    COMPOSES.push(c);
    ACTIVE = c;
    document.getElementById("zone-redaction").appendChild(el);
    ongletRuban = CFG.ongletRedaction || "Message";
    afficherZone();
    noter("ouverture", { compose: mid, type: c.type });
    if (o.focus) { var f = o.focus === "CORPS" ? c.corps : o.focus === "OBJET" ? c.objet : c.puits[o.focus]; if (f) { f.focus(); if (o.focus === "CORPS" || o.focus === "A") placerFinCe(f); } }
    letsignit(c);
    return c;
  }
  function placerFinCe(el) { if (el.tagName === "INPUT") return; var r = document.createRange(); r.selectNodeContents(el); r.collapse(false); var s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  function montrerLigne(c, cle, bascule) {
    c.lignes[cle].hidden = false;
    if (bascule) bascule.remove();
    noter("ligne-affichee", { ligne: cle });
    c.puits[cle].focus();
  }
  function fermerRedaction(c, raison) {
    fermerSuggestions(); fermerMenus();
    c.el.remove();
    COMPOSES.splice(COMPOSES.indexOf(c), 1);
    if (ACTIVE === c) ACTIVE = null;
    noter("fermeture", { compose: c.id, raison: raison });
    afficherZone();
  }
  function afficherZone() {
    var v = composeVisible();
    COMPOSES.forEach(function (c) { c.el.style.display = c === v ? "" : "none"; });
    document.getElementById("lecture").style.display = v ? "none" : "";
    majOngletsDocs();
    rendreRuban();
  }
  function majOngletsDocs() {
    var z = document.getElementById("onglets-docs");
    z.innerHTML = "";
    COMPOSES.forEach(function (c) {
      var d = document.createElement("span"); d.className = "doc";
      var t = document.createElement("span");
      t.setAttribute("role", "tab"); t.setAttribute("aria-selected", String(c === composeVisible())); t.setAttribute("data-ident", "doc:" + c.id); t.tabIndex = 0;
      t.textContent = "Modification de " + ((c.objet && c.objet.value) || "(Aucun objet)");
      t.addEventListener("click", function () { if (ACTIVE) ACTIVE.visible = false; ACTIVE = c; c.visible = true; afficherZone(); });
      d.appendChild(t);
      d.appendChild(bouton("doc:" + c.id + ":fermer", "✕", { aria: "Fermez", action: function () { noter("fermez", {}); fermerRedaction(c, "fermee"); } }));
      z.appendChild(d);
    });
  }
  function repondre(i) {
    var m = MESSAGES[i];
    if (!m) return null;
    var cit = '<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>' +
      '<div id="appendonsend"></div><hr style="display:inline-block;width:98%" tabindex="-1"><div id="divRplyFwdMsg" dir="ltr"><b>De :</b> ' + echapper(m.de) + " &lt;" + echapper(m.adresse) + "&gt;<br><b>Objet :</b> " + echapper(m.objet) + "</div><div>" + echapper(m.corps).replace(/\n/g, "<br>") + "</div>";
    return nouvelleRedaction({ type: "reponse", a: [m.adresse], objet: "RE: " + m.objet, objetVisible: !!CFG.reponseObjetVisible, corpsHtml: cit, focus: "CORPS" });
  }

  // ------------------------------------------------------------ Letsignit
  var SIGNATURE_LETSIGNIT =
    '<table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,sans-serif;font-size:13px"><tbody><tr>' +
    '<td style="padding-right:12px;vertical-align:top"><a href="https://www.ecole-test.fr"><img alt="Logo École Test" width="80" height="80" src="data:image/gif;base64,R0lGODlhAQABAAAAACw="></a></td>' +
    '<td style="border-left:2px solid #c00;padding-left:12px;vertical-align:top">' +
    '<p style="margin:0"><span style="font-weight:bold;color:#c00">Paul DURANDEAU</span></p>' +
    '<p style="margin:0"><span>Chargé de relations entreprises</span></p>' +
    '<p style="margin:0"><span>Campus de Paris</span></p>' +
    '<p style="margin:0"><a href="tel:+33100000000"><span>+33 1 00 00 00 00</span></a></p>' +
    '<p style="margin:4px 0 0"><a href="https://www.linkedin.com/school/ecole-test"><img alt="LinkedIn" width="16" height="16" src="data:image/gif;base64,R0lGODlhAQABAAAAACw="></a> <a href="https://www.facebook.com/ecoletest"><img alt="Facebook" width="16" height="16" src="data:image/gif;base64,R0lGODlhAQABAAAAACw="></a></p>' +
    '</td></tr></tbody></table><p style="margin:6px 0 0"><a href="https://www.ecole-test.fr/journees-portes-ouvertes"><img alt="Bannière portes ouvertes" width="400" height="60" src="data:image/gif;base64,R0lGODlhAQABAAAAACw="></a></p>';
  window.__SIGNATURE_LETSIGNIT = SIGNATURE_LETSIGNIT;
  function letsignit(c) {
    var L = CFG.letsignit;
    if (!L) return;
    if (L.modale) {
      setTimeout(function () {
        if (!c.el.isConnected) return;
        var m = ouvrirModale({ id: "letsignit", titre: "Letsignit Add-in", texte: "Letsignit : mise à jour de votre signature…", piege: L.modale.piege !== false });
        if (L.modale.duree != null) setTimeout(function () { fermerModale(m); }, D(L.modale.duree, L.absolu));
      }, D(L.modale.apres || 0, L.absolu));
    }
    if (L.html === false) return;
    if (L.declencheur === "collage-corps") {
      // Variante : le complément pose (ou repose) sa signature juste après le collage du message.
      c.corps.addEventListener("paste", function () { if (!c.signature) setTimeout(function () { poserSignature(c, L); }, D(L.delai || 0, L.absolu)); });
    } else setTimeout(function () { poserSignature(c, L); }, D(L.delai || 0, L.absolu));
  }
  function poserSignature(c, L) {
    if (!c.el.isConnected) return;
    var html = L.html || SIGNATURE_LETSIGNIT;
    if (L.conteneur === null) {
      var w = document.createElement("div"); w.innerHTML = html;
      while (w.firstChild) c.corps.appendChild(w.firstChild);
    } else {
      var s = c.corps.querySelector("#Signature");
      if (!s) { s = document.createElement("div"); s.id = "Signature"; s.className = "elementToProof"; c.corps.appendChild(s); }
      s.innerHTML = html;
    }
    c.signature = true;
    noter("signature", { compose: c.id });
  }

  // ------------------------------------------------------------ envoi
  function destinataires(c) {
    var out = { a: [], cc: [], cci: [], invalides: [], libres: [] };
    [["A", "a"], ["CC", "cc"], ["CCI", "cci"]].forEach(function (x) {
      var champ = c.puits[x[0]];
      infosChamp(champ).forEach(function (i) { if (i.resolu && i.adresse) out[x[1]].push(i.adresse); else out.invalides.push(i.texte || "?"); });
      var libre = texteLibre(champ).trim();
      if (libre) {
        // Comme Outlook à l'envoi : un texte resté tapé est résolu (adresse complète) ou refusé.
        libre.split(/[;,]/).map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (s) {
          var r = resoudre(s);
          out.libres.push(s);
          if (r.resolu) out[x[1]].push(r.adresse); else out.invalides.push(s);
        });
      }
    });
    return out;
  }
  function cliquerEnvoyer(c) {
    if (!c.el.isConnected) return;
    var dest = destinataires(c);
    noter("envoyer-clique", { compose: c.id });
    if (dest.invalides.length) {
      ouvrirModale({ id: "verifier-noms", alerte: true, titre: "Vérifiez les noms", texte: "Outlook ne reconnaît pas : " + dest.invalides.join(", "), boutons: [{ texte: "OK", ident: "dialogue:verifier-noms:ok" }] });
      return;
    }
    var tous = dest.a.concat(dest.cc, dest.cci);
    if (!tous.length) {
      ouvrirModale({ id: "sans-destinataire", alerte: true, titre: "Ajoutez un destinataire", texte: "Ce message doit comporter au moins un destinataire.", boutons: [{ texte: "OK", ident: "dialogue:sans-destinataire:ok" }] });
      return;
    }
    if (CFG.envoiDialogue === "sensibilite" && !c.etiquette) {
      ouvrirModale({ id: "sensibilite", titre: "Étiquette de sensibilité requise", texte: "Votre organisation exige une étiquette de sensibilité pour envoyer ce message. Choisissez-en une, puis envoyez.",
        boutons: [{ texte: "OK", ident: "dialogue:sensibilite:ok", action: function () { c.etiquette = "Général"; } }, { texte: "Annuler", ident: "dialogue:sensibilite:annuler" }] });
      return;
    }
    if (CFG.envoiDialogue === "externes" && !c.externesOk && tous.some(estExterne)) {
      ouvrirModale({ id: "externes", titre: "Destinataires externes", texte: "Ce message sera envoyé à des personnes extérieures à votre organisation.",
        boutons: [{ texte: "Envoyer quand même", ident: "dialogue:externes:envoyer", action: function () { c.externesOk = true; cliquerEnvoyer(c); } }, { texte: "Ne pas envoyer", ident: "dialogue:externes:annuler" }] });
      return;
    }
    var rec = { a: dest.a, cc: dest.cc, cci: dest.cci, objet: c.objet ? c.objet.value : "", corps: c.corps.innerText, texte: c.corps.textContent, html: c.corps.innerHTML,
      importance: c.importance, type: c.type, libres: dest.libres, t: Date.now(), compose: c.id };
    if (!navigator.onLine) {
      if (CFG.horsLigne === "echec") {
        c.el.querySelector(".bandeau-redaction").innerHTML = '<div role="alert" class="bandeau">Impossible d\'envoyer ce message. Vérifiez votre connexion.</div>';
        noter("envoi-echec", { compose: c.id });
        return;
      }
      BOITE.push(rec); publier();
      noter("boite-envoi", { compose: c.id });
      fermerRedaction(c, "boite-envoi");
      return;
    }
    var finir = function () {
      if (!c.el.isConnected) return;
      envoyes.push(rec); publier();
      noter("envoye", { compose: c.id, a: dest.a, cci: dest.cci.length });
      fermerRedaction(c, "envoyee");
    };
    if (CFG.envoiLent) { c.el.querySelector(".bandeau-redaction").innerHTML = '<div role="status" class="bandeau">Envoi en cours…</div>'; setTimeout(finir, D(CFG.envoiLent)); }
    else finir();
  }
  function bandeauHorsLigne() {
    var b = document.getElementById("bandeaux");
    b.innerHTML = navigator.onLine ? "" : '<div role="status" class="bandeau" data-ident="hors-ligne">Vous êtes hors connexion. Vos messages seront envoyés dès que la connexion sera rétablie.</div>';
  }
  window.addEventListener("offline", function () { noter("hors-ligne", {}); bandeauHorsLigne(); });
  window.addEventListener("online", function () {
    noter("en-ligne", { boite: BOITE.length });
    bandeauHorsLigne();
    // Outlook vide la Boîte d'envoi d'un coup au retour de la connexion.
    BOITE.splice(0).forEach(function (r) { r.depuisBoiteEnvoi = true; r.tEnvoi = Date.now(); envoyes.push(r); });
    publier();
  });

  // ------------------------------------------------------------ recherche
  var rech = document.getElementById("recherche"), rs = document.getElementById("recherche-sugg");
  rech.addEventListener("input", function () {
    var q = rech.value.trim();
    rs.innerHTML = "";
    if (!q) { rs.hidden = true; return; }
    chercher(q).slice(0, 4).forEach(function (c) { var o = document.createElement("div"); o.setAttribute("role", "option"); o.setAttribute("data-ident", "recherche:personne"); o.textContent = c.nom + " " + c.adresse; rs.appendChild(o); });
    var o = document.createElement("div"); o.setAttribute("role", "option"); o.textContent = "Rechercher « " + q + " »"; rs.appendChild(o);
    rs.hidden = false;
  });
  rech.addEventListener("blur", function () { setTimeout(function () { rs.hidden = true; }, 0); });

  // ------------------------------------------------------------ crochets pour les tests
  window.__fausse = {
    etat: function () {
      return {
        composes: COMPOSES.map(function (c) {
          return { id: c.id, type: c.type, visible: c === composeVisible(), a: infosChamp(c.puits.A), cc: infosChamp(c.puits.CC), cci: infosChamp(c.puits.CCI),
            libres: { A: texteLibre(c.puits.A), CC: texteLibre(c.puits.CC), CCI: texteLibre(c.puits.CCI) },
            objet: c.objet ? c.objet.value : null, corps: c.corps.textContent, html: c.corps.innerHTML, importance: c.importance, cciVisible: !c.lignes.CCI.hidden };
        }),
        envoyes: envoyes.length, boite: BOITE.length, carnet: window.__carnet || 0, modales: MODALES.map(function (m) { return m.o.id; }), menus: MENUS.length,
        actif: ident(document.activeElement), onglet: ongletRuban,
      };
    },
    modale: function (o) { return ouvrirModale(o) && true; },
    fermerModales: function () { MODALES.slice().forEach(fermerModale); },
    sessionExpiree: function () { ouvrirModale({ id: "session", titre: "Session expirée", texte: "Votre session a expiré. Actualisez la page pour vous reconnecter.", boutons: [{ texte: "Actualiser", ident: "dialogue:session:actualiser", action: function () { location.reload(); } }] }); },
    redirigerConnexion: function () { location.href = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=outlook&redirect_uri=https%3A%2F%2Foutlook.cloud.microsoft%2Fmail%2F"; },
    repondre: function (i) { return !!repondre(i); },
    nouveau: function (focus) { return !!nouvelleRedaction({ type: "nouveau", focus: focus || "A" }); },
    selectionner: selectionner,
    poserSignatureMaintenant: function () { if (ACTIVE) poserSignature(ACTIVE, CFG.letsignit || {}); },
    pastille: function (cle, adresse) { var c = composeVisible(); if (c) ajouterPastilles(c.puits[cle], [resoudre(adresse)], "test"); },
  };

  // ------------------------------------------------------------ démarrage
  rendreListe();
  rendreLecture();
  rendreRuban();
  if (CFG.brouillon) {
    var b = CFG.brouillon;
    nouvelleRedaction({ type: "brouillon", a: b.a || [], objet: b.objet || "", corpsHtml: b.corpsHtml || ('<div class="elementToProof" style="font-family: Aptos, sans-serif; font-size: 12pt;">' + echapper(b.corps || "").replace(/\n/g, "<br>") + "</div>"), focus: b.focus || null });
  }
  if (CFG.reponse) repondre(CFG.reponse.depuis != null ? CFG.reponse.depuis : 1);
  noter("pret", {});
})();
