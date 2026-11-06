// Lecture du menu de suggestions -> lignes pour Moteur.parseRows.
//
// Seul morceau du script de contenu éprouvé hors navigateur (jsdom) :
// tests/unit/lecture.test.mjs lui donne de vrais fragments de menus OWA/Gmail
// et vérifie que parseRows en tire les bons { name, mail }.
//
// Règle d'or : on garde les CELLULES SÉPARÉES (un tableau de textes par
// suggestion), jamais recollées. Coller « WANG » + « g.wang@x.fr » donnerait la
// mauvaise adresse « wangg.wang@x.fr » à la regex de parseRows ; « WANG » puis
// « g.wang@x.fr » séparés donnent « g.wang@x.fr ». On lit textContent (jamais
// innerText, indéfini sous jsdom et qui recolle les cellules voisines).
//
// Même emballage UMD que planif.js / moteur.js : require() dans les tests,
// globalThis.AOLecture dans le script de contenu.

(function (root) {
  "use strict";

  // Tables de sélecteurs, par messagerie. Hypothèses à confiance variable ;
  // le Diagnostic (lancé par l'utilisateur sur sa vraie page) les confirme.
  // Premier sélecteur visible gagnant.
  var SEL = {
    gmail: {
      // Bouton « Nouveau message » : attribut d'action stable, puis classe
      // (obfusquée, susceptible de changer), puis repli par libellé.
      compose: ['div[gh="cm"]', '.T-I.T-I-KE.L3'],
      composeText: ['compose', 'nouveau message'],
      // Champ « À » : l'input peoplekit actuel, puis des replis génériques.
      toField: ['div[name="to"] input[peoplekit-id]', 'div[name="to"] input',
        'input[aria-label="To recipients"]', '[role="combobox"][aria-autocomplete="list"]',
        'textarea[name="to"]'],
      chips: 'div[name="to"] [data-hovercard-id]',
      option: '[role="option"]'
    },
    outlook: {
      compose: ['[aria-label="Nouveau message"]', '[aria-label="New mail"]', '[aria-label="New message"]', '[aria-label="Nouveau courrier"]', '[aria-label="Nouveau courriel"]', '[aria-label="New email"]', '[aria-label="Neue E-Mail"]', '[aria-label="Neue Nachricht"]', '[aria-label="Correo nuevo"]', '[aria-label="Mensaje nuevo"]', '[aria-label="Nuova e-mail"]', '[aria-label="Nuovo messaggio"]', '[aria-label="Nieuwe e-mail"]', '[aria-label="Nieuw bericht"]', '[aria-label="Novo email"]', '[aria-label="Nova mensagem"]'],
      composeText: ['nouveau message', 'new mail', 'new message', 'nouveau courrier', 'nouveau courriel', 'new email', 'neue e-mail', 'neue nachricht', 'correo nuevo', 'mensaje nuevo', 'nuova e-mail', 'nuovo messaggio', 'nieuwe e-mail', 'nieuw bericht', 'novo email', 'nova mensagem'],
      toField: ['[aria-label="To"] input', 'input[aria-label="To"]', 'input[aria-label="À"]',
        'input[aria-label="Destinataires"]', '[role="textbox"][aria-label*="To"]', '.ms-BasePicker-input',
        'div[contenteditable="true"][aria-label="À"]', 'div[contenteditable="true"][aria-label="To"]', 'div[contenteditable="true"][aria-label="An"]', 'div[contenteditable="true"][aria-label="Para"]', 'div[contenteditable="true"][aria-label="A"]', 'div[contenteditable="true"][aria-label="Aan"]', 'input[aria-label="An"]', 'input[aria-label="Para"]', 'input[aria-label="A"]', 'input[aria-label="Aan"]'],
      chips: '.ms-PickerPersona-container, [data-selection-index], [role="listitem"], [class*="_EType_RECIPIENT_ENTITY"], [id^="REK"]',
      option: '[role="option"]',
      name: '.ms-Persona-primaryText',
      mail: '.ms-Persona-secondaryText'
    }
  };

  // Une seule adresse, pour isoler l'adresse dans un texte. Domaine ASCII
  // (comme moteur.js : name@école.fr n'est volontairement pas reconnu).
  // Bornée des deux côtés : jamais un morceau d'adresse (« andré.martin@ » ne donne
  // pas « martin@ », une autre personne).
  var MAIL = /(?<=^|[\s;,<>()\[\]"':=|])[A-Za-z0-9._%+'\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}(?=$|[\s;,<>()\[\]"':|.?!])/u;

  function txt(el) { return el && typeof el.textContent === "string" ? el.textContent.trim() : ""; }

  // Les éléments-feuilles (sans enfant élément) : leur texte ne recolle pas
  // celui des voisins, contrairement à un conteneur.
  // Balises de mise en forme EN LIGNE (partie tapée en gras, surlignage…) : elles
  // ne découpent pas une cellule. « <b>ma</b>rtin.dupont@… » reste une seule cellule.
  var FORMAT = /^(B|STRONG|EM|I|U|MARK|SMALL|SUB|SUP|S|CODE|FONT|WBR)$/;
  function feuilleLogique(e) {
    for (var c = e.firstElementChild; c; c = c.nextElementSibling) if (!FORMAT.test(c.tagName) || !feuilleLogique(c)) return false;
    return true;
  }
  function leafTexts(el) {
    var out = [], seen = {};
    var all = [].slice.call(el.querySelectorAll("*"));
    for (var i = 0; i < all.length; i++) {
      var e = all[i];
      if (!feuilleLogique(e)) continue;
      if (e.parentElement && e.parentElement !== el && el.contains(e.parentElement) && feuilleLogique(e.parentElement)) continue; // déjà pris avec son parent
      var t = txt(e);
      if (t && !seen[t]) { seen[t] = 1; out.push(t); }
    }
    // Un nom (ou une adresse) posé en texte nu À CÔTÉ d'un élément (ex.
    // « Jean Dupont <span>adresse</span> ») : on le prend comme cellule à part,
    // pour ne jamais le recoller à l'adresse.
    try {
      var doc = el.ownerDocument || document;
      var w = doc.createTreeWalker(el, 4 /* NodeFilter.SHOW_TEXT */, null, false), n;
      while ((n = w.nextNode())) {
        var p = n.parentNode;
        if (!p || !p.children || p.children.length === 0) continue; // déjà pris comme élément-feuille
        if (p.nodeType === 1 && feuilleLogique(p)) continue; // texte d'une cellule mise en forme : déjà pris entier
        var s = (n.data || "").trim();
        if (s && !seen[s]) { seen[s] = 1; out.push(s); }
      }
    } catch (e) {}
    // Certaines versions ne mettent l'adresse que dans title/aria-label : on la
    // prend comme cellule séparée (uniquement si c'est bien une adresse).
    try {
      [].slice.call(el.querySelectorAll("[title],[aria-label]")).concat([el]).forEach(function (e) {
        ["title", "aria-label"].forEach(function (att) {
          var v = ((e.getAttribute && e.getAttribute(att)) || "").trim();
          if (v && v.indexOf("@") >= 0 && MAIL.test(v) && !seen[v]) { seen[v] = 1; out.push(v); }
        });
      });
    } catch (e2) {}
    return out;
  }

  // Repli commun : les cellules-feuilles, SÉPARÉES (jamais recollées). S'il n'y a
  // pas de nom en élément-feuille (nom en texte nu à côté de l'adresse), on rend
  // le bloc entier « Nom adresse » que parseRows sait découper sans recoller.
  var MAIL_G = new RegExp(MAIL.source, "gu");
  // Badges qui ne sont pas un nom (« Externe » sur les contacts hors organisation).
  var BADGE = /^(externe|external|invit[ée]|guest|hors organisation)$/i;
  var GROUPE = /^(groupe|group|groupe microsoft 365|microsoft 365 group|liste de distribution|distribution list|liste de diffusion|liste de contacts|contact list|salle|room|salle de r[ée]union|ressource|resource|[ée]quipement|equipment)$/i;
  var LOCALE_GROUPE = /^(salle|salles|room|rooms|amphi|tous|toutes|all|liste|list|listes|promo|promos|groupe|groupes|group|groups|dl|diffusion|everyone)([._-]|\d|$)/i;
  function leafCells(opt) {
    var leaves = leafTexts(opt);
    // Les adresses DISTINCTES portées par les feuilles. La même adresse peut
    // apparaître plusieurs fois (texte + title + aria-label de l'option, comme
    // dans le nouvel Outlook web) : ce n'est pas une ligne à deux adresses.
    var seen = {}, distinct = [];
    for (var i = 0; i < leaves.length; i++) {
      var found = leaves[i].match(MAIL_G) || [];
      for (var j = 0; j < found.length; j++) { var k = found[j].toLowerCase(); if (!seen[k]) { seen[k] = found[j]; distinct.push(k); } }
    }
    if (distinct.length === 1) {
      var mail = seen[distinct[0]];
      // Groupe, liste de diffusion ou salle : pas une personne, on l'écarte.
      if (leaves.some(function (s) { return GROUPE.test(s); }) || LOCALE_GROUPE.test(mail.split("@")[0])) return null;
      // Le nom : la plus longue feuille sans adresse, hors badge (« Externe ») et
      // hors initiales d'avatar (« AP »). Jamais recollé à l'adresse.
      var names = leaves.filter(function (s) { return !MAIL.test(s) && !BADGE.test(s) && !/^[A-Z]{1,3}$/.test(s); })
        .sort(function (a, b) { return b.length - a.length; });
      return names.length ? [names[0], mail] : [mail];
    }
    // Deux adresses distinctes : ligne ambiguë (parasite), on l'écarte.
    if (distinct.length > 1) return null;
    // Aucune feuille à adresse : le bloc entier, « Nom - adresse » d'une pièce
    // (séparateur présent), que parseRows sait découper. JAMAIS quand des
    // feuilles existent sans adresse : textContent recollerait nom et adresse.
    var block = txt(opt);
    return MAIL.test(block) ? [block] : null;
  }

  // Une option -> ses cellules [texte, ...] (nom(s) puis adresse), ou null si
  // aucune adresse exploitable.
  function optionCells(opt, source) {
    if (source === "gmail") {
      // Adresse fiable = data-hovercard-id, porté par l'option elle-même ou un
      // descendant. Avec lui : nom = texte sans l'adresse. SANS lui, on NE recolle
      // PAS (sinon « nomadresse ») : on passe par les cellules-feuilles séparées.
      var hid = (opt.getAttribute("data-hovercard-id") || "").trim();
      if (!hid || hid.indexOf("@") < 0) {
        var holder = opt.querySelector("[data-hovercard-id]");
        hid = holder ? (holder.getAttribute("data-hovercard-id") || "").trim() : "";
      }
      if (hid && hid.indexOf("@") >= 0) {
        var all = txt(opt), idx = all.toLowerCase().indexOf(hid.toLowerCase());
        // On n'enlève que les séparateurs OUVRANTS/espaces en fin (pas « ) » ni
        // « > », qui peuvent terminer un vrai nom « Marie (Marketing) »).
        var name = (idx >= 0 ? all.slice(0, idx) : all.replace(MAIL, " ")).replace(/[\s<(·,–\-]+$/, "").trim();
        return name ? [name, hid] : [hid, hid]; // sans nom : adresse répétée (parseRows le gère)
      }
      return leafCells(opt);
    }
    // Outlook : nom + adresse en deux cellules « persona », sinon repli feuilles.
    var primary = opt.querySelector(SEL.outlook.name);
    var secondary = opt.querySelector(SEL.outlook.mail);
    if (secondary && MAIL.test(txt(secondary))) {
      var cells = [];
      if (primary && txt(primary)) cells.push(txt(primary));
      cells.push(txt(secondary));
      return cells;
    }
    return leafCells(opt);
  }

  // scope = la liste de suggestions résolue (ou document). source = 'gmail'|'outlook'.
  function lireRows(scope, source) {
    if (!scope) return [];
    var sel = source === "gmail" ? SEL.gmail : SEL.outlook;
    var opts = [].slice.call(scope.querySelectorAll(sel.option));
    if (!opts.length) {
      // Replis génériques quand la liste n'utilise pas [role=option] : éléments de
      // liste d'une [role=listbox], ou enfants directs contenant une adresse.
      opts = [].slice.call(scope.querySelectorAll('[role="listbox"] li,[aria-selected]'));
      if (!opts.length && scope.children) opts = [].slice.call(scope.children).filter(function (c) { return MAIL.test(c.textContent || ""); });
    }
    var rows = [];
    for (var i = 0; i < opts.length; i++) {
      var cells = optionCells(opts[i], source);
      if (cells && cells.length) rows.push(cells);
    }
    return rows;
  }

  var out = { lireRows: lireRows, SEL: SEL, MAIL: MAIL };
  if (typeof module !== "undefined" && module.exports) module.exports = out;
  else root.AOLecture = out;
})(typeof globalThis !== "undefined" ? globalThis : this);
