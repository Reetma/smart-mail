// Adaptateur DOM : les gestes que le moteur demande, réalisés dans la page de
// la messagerie (champ « À » d'un nouveau message). Ne tape QUE des lettres,
// des chiffres et un espace intérieur, jamais Entrée/Tab/virgule : rien n'est
// envoyé. S'il voit qu'un destinataire a été validé par erreur (une pastille
// apparaît), focus() rend { ok:false } et le moteur s'arrête aussitôt.
//
// Tout est exposé via window.AODom (les fichiers suivants du script de contenu
// le lisent). S'appuie sur AOLecture (sélecteurs + lecture des lignes).

(function (root) {
  "use strict";

  var SEL = (root.AOLecture || {}).SEL || { gmail: {}, outlook: {} };
  var lireRows = (root.AOLecture || {}).lireRows || function () { return []; };

  // Messagerie : forcée par les fausses pages de test, sinon déduite de l'URL.
  var SOURCE = (function () {
    try {
      var d = document.documentElement.dataset.aoSource;
      if (d === "gmail" || d === "outlook") return d;
    } catch (e) {}
    return /mail\.google\.com/.test(location.hostname) ? "gmail" : "outlook";
  })();

  // En test, toutes les pauses tombent à ~1 ms (sinon le rythme réel).
  var AO_WAIT = (root.__AO_TEST && root.__AO_TEST.waitMs != null) ? root.__AO_TEST.waitMs : null;
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, AO_WAIT != null ? AO_WAIT : ms); }); }

  function isField(el) { return el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA") && typeof el.value === "string"; }
  // offsetParent est null pour un élément position:fixed (menu « portalisé ») même
  // visible : on accepte aussi getClientRects() non vide. aria-hidden exclut.
  function visible(el) {
    if (!el) return false;
    if (el.getAttribute && el.getAttribute("aria-hidden") === "true") return false;
    if (!(el.offsetParent !== null || (el.getClientRects && el.getClientRects().length > 0))) return false;
    try { var cs = window.getComputedStyle(el); if (cs.visibility === "hidden" || cs.display === "none" || parseFloat(cs.opacity) === 0) return false; } catch (e) {}
    return true;
  }
  // Un champ hors d'une fenêtre de rédaction ne doit pas être pris pour le champ
  // « À » : on écarte la barre de recherche (landmark search/banner/header, ou
  // libellé « search »/« recherche »).
  function notSearch(el) {
    if (!el) return false;
    if (el.closest && el.closest('[role="search"],[role="banner"],header')) return false;
    var al = ((el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("placeholder"))) || "").toLowerCase();
    return !/search|recherch/.test(al);
  }

  // ----------------------------------------------------------- frappe
  function keyInfo(t) {
    if (t === "Backspace") return { key: "Backspace", code: "Backspace", keyCode: 8 };
    if (t === " ") return { key: " ", code: "Space", keyCode: 32 };
    if (t >= "a" && t <= "z") return { key: t, code: "Key" + t.toUpperCase(), keyCode: t.toUpperCase().charCodeAt(0) };
    if (t >= "0" && t <= "9") return { key: t, code: "Digit" + t, keyCode: t.charCodeAt(0) };
    return { key: t, code: "", keyCode: t.charCodeAt(0) || 0 };
  }
  function fireKey(el, type, t) {
    var k = keyInfo(t), ev = new KeyboardEvent(type, { key: k.key, code: k.code, bubbles: true, cancelable: true, composed: true });
    try { Object.defineProperty(ev, "keyCode", { get: function () { return k.keyCode; } }); Object.defineProperty(ev, "which", { get: function () { return k.keyCode; } }); } catch (e) {}
    return el.dispatchEvent(ev);
  }
  function setNativeValue(el, v) {
    var proto = el.tagName === "TEXTAREA" ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value").set.call(el, v);
  }
  function fireInput(el, inputType, data) {
    el.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, composed: true, inputType: inputType, data: data }));
    el.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: false, composed: true, inputType: inputType, data: data }));
  }
  function ensureCaret(el) {
    if (isField(el)) return;
    var s = window.getSelection(), inside = s && s.rangeCount && el.contains(s.anchorNode);
    if (!inside) { var r = document.createRange(); r.selectNodeContents(el); r.collapse(false); s.removeAllRanges(); s.addRange(r); }
  }
  // Le curseur est-il (encore) dans le champ visé ? Jamais de frappe ni
  // d'effacement ailleurs : corps d'un message, objet, recherche, discussion…
  function focusDans(el) { var a = document.activeElement; return !!a && (a === el || (el.contains && el.contains(a))); }
  // Texte libre du champ (hors pastilles de destinataires et caractères invisibles).
  function texteLibre(el) {
    if (isField(el)) return el.value;
    var out = "", w = document.createTreeWalker(el, 4); // NodeFilter.SHOW_TEXT
    for (var n = w.nextNode(); n; n = w.nextNode()) {
      var p = n.parentElement && n.parentElement.closest('[contenteditable="false"]');
      if (p && p !== el && el.contains(p)) continue; // texte d'une pastille
      out += n.nodeValue;
    }
    return out.replace(/[\u200B-\u200D\u2060\uFEFF]/g, "");
  }
  // Curseur à la toute fin du champ (là où Smart Mail a tapé).
  function caretFin(el) {
    if (isField(el)) { try { el.selectionStart = el.selectionEnd = el.value.length; } catch (e) {} return; }
    var s = window.getSelection(), r = document.createRange(); r.selectNodeContents(el); r.collapse(false); s.removeAllRanges(); s.addRange(r);
  }
  // Ce que Smart Mail a tapé dans le champ et pas encore effacé.
  var tape = "";
  // Arrêt sans le curseur (l'utilisateur est ailleurs) : on retire QUAND MÊME nos
  // lettres du champ où on les a tapées (ou de celui qui l'a remplacé), sans lui
  // prendre le curseur, et seulement si le champ se termine bien par elles.
  function nettoyerSansFocus() {
    if (!tape) return { ok: true };
    var el = (locked && locked.isConnected) ? locked : remplacant();
    var a = root.AOArmement;
    if (!el && a && a.field && a.field() && a.field().isConnected) el = a.field();
    if (!el) return { ok: false };
    var libre = texteLibre(el);
    if (libre.slice(-tape.length).toLowerCase() !== tape.toLowerCase()) return { ok: false };
    if (isField(el)) {
      setNativeValue(el, el.value.slice(0, el.value.length - tape.length));
      fireInput(el, "deleteContentBackward", null);
    } else {
      var reste = tape.length, noeuds = [], w = document.createTreeWalker(el, 4);
      for (var n = w.nextNode(); n; n = w.nextNode()) {
        var p = n.parentElement && n.parentElement.closest('[contenteditable="false"]');
        if (!(p && p !== el && el.contains(p))) noeuds.push(n);
      }
      for (var i = noeuds.length - 1; i >= 0 && reste > 0; i--) {
        var v = noeuds[i].nodeValue, garde = v.replace(/[\u200B-\u200D\u2060\uFEFF]+$/, "");
        var k = Math.min(reste, garde.length);
        noeuds[i].nodeValue = garde.slice(0, garde.length - k) + v.slice(garde.length);
        reste -= k;
      }
      fireInput(el, "deleteContentBackward", null);
    }
    var ok = texteLibre(el).slice(-tape.length).toLowerCase() !== tape.toLowerCase() || !texteLibre(el).replace(/\s/g, "");
    if (ok) tape = "";
    return { ok: ok };
  }
  function insertText(el, text) {
    ensureCaret(el);
    var ok = false; try { ok = document.execCommand("insertText", false, text); } catch (e) { ok = false; }
    if (ok) return true;
    if (isField(el)) {
      var s = el.selectionStart == null ? el.value.length : el.selectionStart, e2 = el.selectionEnd == null ? el.value.length : el.selectionEnd;
      setNativeValue(el, el.value.slice(0, s) + text + el.value.slice(e2));
      try { el.selectionStart = el.selectionEnd = s + text.length; } catch (e) {}
      fireInput(el, "insertText", text); return true;
    }
    var sel = window.getSelection();
    if (sel && sel.rangeCount) { var r = sel.getRangeAt(0); r.deleteContents(); var node = document.createTextNode(text); r.insertNode(node); r.setStartAfter(node); r.setEndAfter(node); sel.removeAllRanges(); sel.addRange(r); }
    else el.textContent += text;
    fireInput(el, "insertText", text); return true;
  }
  // Tape lettre par lettre, en revérifiant AVANT CHAQUE LETTRE que le curseur est
  // toujours dans le champ (sinon arrêt : { ok:false }, rien n'est tapé ailleurs).
  async function taper(el, text) {
    var chars = String(text);
    for (var i = 0; i < chars.length; i++) {
      if (!focusDans(el)) return { ok: false, raison: "focus" };
      var ch = chars[i];
      fireKey(el, "keydown", ch); fireKey(el, "keypress", ch);
      if (!focusDans(el)) return { ok: false, raison: "focus" }; // la page a déplacé le curseur pendant la touche
      caretFin(el);
      insertText(el, ch);
      tape += ch;
      fireKey(el, "keyup", ch); await sleep(35);
    }
    return { ok: true };
  }
  // N'efface QUE ce que Smart Mail a tapé, lettre par lettre, et seulement si le
  // champ se termine toujours par ce texte et que le curseur y est encore : sinon
  // { ok:false } et rien n'est effacé (jamais une pastille, jamais le texte de
  // l'utilisateur, jamais un autre champ).
  async function effacer(el, n) {
    n = Math.min(n, tape.length);
    for (var i = 0; i < n; i++) {
      if (!focusDans(el)) return { ok: false, raison: "focus" };
      var libre = texteLibre(el);
      if (!tape || libre.slice(-tape.length).toLowerCase() !== tape.toLowerCase()) return { ok: false, raison: "modifie" };
      caretFin(el);
      fireKey(el, "keydown", "Backspace");
      if (!focusDans(el)) return { ok: false, raison: "focus" }; // la page a déplacé le curseur pendant la touche
      var ok = false; try { ok = document.execCommand("delete", false); } catch (e) { ok = false; }
      if (!ok && isField(el)) { var v = el.value; setNativeValue(el, v.slice(0, Math.max(0, v.length - 1))); try { el.selectionStart = el.selectionEnd = el.value.length; } catch (e) {} fireInput(el, "deleteContentBackward", null); }
      else if (!ok) { var s = window.getSelection(); if (s) { s.modify("extend", "backward", "character"); document.execCommand("delete"); } fireInput(el, "deleteContentBackward", null); }
      fireKey(el, "keyup", "Backspace");
      if (texteLibre(el).length >= libre.length) return { ok: false, raison: "effacement" }; // rien n'a été effacé : on n'insiste pas
      tape = tape.slice(0, -1);
      await sleep(25);
    }
    return { ok: true };
  }

  // ------------------------------------------------- champ, focus, liste
  function fieldSelectors() { return (SOURCE === "gmail" ? SEL.gmail : SEL.outlook).toField || []; }
  function findField() {
    var sels = fieldSelectors();
    var dialogs = [].slice.call(document.querySelectorAll('[role="dialog"]')).reverse(); // le plus récent d'abord
    for (var d = 0; d < dialogs.length; d++) for (var s = 0; s < sels.length; s++) { var el = dialogs[d].querySelector(sels[s]); if (visible(el)) return el; }
    // Repli hors dialogue : on écarte la barre de recherche (sinon on taperait dedans).
    for (var s2 = 0; s2 < sels.length; s2++) { var el2 = document.querySelector(sels[s2]); if (visible(el2) && notSearch(el2)) return el2; }
    return null;
  }
  var locked = null, ancetres = [], libelleChamp = "";
  // Outlook peut remplacer le champ « À » par un nouvel élément (réaffichage) : on
  // reprend celui qui a le même libellé au même endroit (ancêtres encore là).
  function remplacant() {
    for (var i = 0; i < ancetres.length; i++) {
      var p = ancetres[i];
      if (!p || !p.isConnected) continue;
      var c = [].slice.call(p.querySelectorAll('input,textarea,[contenteditable="true"]')).filter(function (e) {
        return (e.getAttribute("aria-label") || "") === libelleChamp && visible(e);
      });
      return c.length === 1 ? c[0] : null;
    }
    return null;
  }
  function memoriser(el) {
    ancetres = []; libelleChamp = (el && el.getAttribute && el.getAttribute("aria-label")) || "";
    for (var p = el && el.parentElement, i = 0; p && i < 4 && p !== document.body; p = p.parentElement, i++) ancetres.push(p);
  }
  function currentField() {
    var a = root.AOArmement;
    if (a && a.isArmed()) return a.field(); // le champ armé par l'utilisateur prime
    if (locked && document.contains(locked) && visible(locked)) return locked;
    if (locked && !document.contains(locked)) { var r = remplacant(); if (r) { locked = r; return r; } }
    return findField();
  }
  // Le curseur peut être remis dans le champ seulement s'il n'est nulle part
  // (page), déjà dans le champ, ou dans la liste de suggestions / la ligne
  // « plus de résultats » (nos propres clics). S'il est ailleurs (l'utilisateur
  // a cliqué ou tape autre part), on ne le reprend PAS : { ok:false }, arrêt.
  function peutReprendre(el) {
    var a = document.activeElement;
    if (!a || a === document.body || a === document.documentElement || a === el || (el.contains && el.contains(a))) return true;
    if (a.closest && a.closest('[role="listbox"],[role="option"]')) return true;
    return !!(a.textContent && moreLabel(a));
  }
  function focus(verrouiller) {
    var el = currentField();
    if (!el) return { ok: false };
    if (!peutReprendre(el)) return { ok: false, raison: "ailleurs" };
    el.focus(); if (verrouiller) { locked = el; memoriser(el); }
    var ok = document.activeElement === el && visible(el);
    if (ok && newChipAppeared()) return { ok: false, raison: "pastille" }; // un destinataire a été validé : arrêt
    return { ok: ok };
  }
  // La boîte de réception est elle-même une [role="listbox"] permanente
  // (« Liste des messages ») : il ne faut JAMAIS la lire comme des suggestions
  // (en plus, ses aperçus contiennent parfois une adresse). On l'écarte par son
  // libellé, ainsi que notre propre panneau.
  var MAIL_RE = /[A-Za-z0-9._%+'\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}/;
  function boxHasMail(el) { return !!el && MAIL_RE.test(el.textContent || ""); }
  function inPanel(el) { return !!(el && el.closest && el.closest('[data-testid="ao-panel"]')); }
  function isMsgList(box) {
    var al = ((box && box.getAttribute && box.getAttribute("aria-label")) || "").toLowerCase();
    return /liste des messages|liste de messages|message list|courrier|bo[iî]te de r[ée]ception|inbox|nachrichtenliste|posteingang|lista de mensajes|bandeja de entrada|elenco messaggi|posta in arrivo|berichtenlijst|postvak in|lista de mensagens|caixa de entrada/.test(al);
  }
  function usable(box) { return visible(box) && !isMsgList(box) && !inPanel(box); }
  // Résultats d'une recherche de courrier (jamais des suggestions du champ « À »).
  function isResults(box) {
    var al = ((box.getAttribute && box.getAttribute("aria-label")) || "").toLowerCase();
    return /r[ée]sultats?|results?|recherche|search/.test(al) || !!(box.closest && box.closest('[role="main"],[role="search"]'));
  }

  // La VRAIE liste de suggestions liée au champ : d'abord par ARIA (si OWA la
  // remplit en tapant), sinon une [role="listbox"] à adresses qui n'est ni la
  // liste des messages ni notre panneau, sinon le conteneur d'une option à
  // adresse (hors liste des messages). Jamais de repli sur « tout le document ».
  function resolveListbox(el) {
    var ids = el && el.getAttribute && (el.getAttribute("aria-controls") || el.getAttribute("aria-owns"));
    if (ids) { var parts = ids.split(/\s+/); for (var i = 0; i < parts.length; i++) { var lb = document.getElementById(parts[i]); if (lb && usable(lb)) return lb; } }
    var ad = el && el.getAttribute && el.getAttribute("aria-activedescendant");
    if (ad) { var opt = document.getElementById(ad), lb2 = opt && opt.closest('[role="listbox"]'); if (lb2 && usable(lb2)) return lb2; }
    // Repli : une liste à adresses, mais seulement sous le champ (jamais les
    // résultats d'une recherche, la liste des messages ou un volet ailleurs).
    var boxes = [].slice.call(document.querySelectorAll('[role="listbox"]')).filter(function (b) { return usable(b) && boxHasMail(b) && !isResults(b) && nearField(b) && listesInitiales.indexOf(b) < 0; });
    if (boxes.length) return boxes[0];
    var opts = [].slice.call(document.querySelectorAll('[role="option"]')).filter(function (o) {
      if (!visible(o) || inPanel(o) || !MAIL_RE.test(o.textContent || "") || !nearField(o)) return false;
      if (o.closest && o.closest('[role="main"],[role="search"],[role="grid"],[role="tree"]')) return false;
      var lb = o.closest && o.closest('[role="listbox"]');
      if (lb && listesInitiales.indexOf(lb) >= 0) return false;
      var own = o.closest('[role="listbox"]');
      return !(own && isMsgList(own)); // jamais une option de la boîte de réception
    });
    if (opts.length) return opts[0].closest('[role="listbox"]') || opts[0].parentNode;
    return null;
  }
  // OWA n'affiche que ~5 suggestions, puis une ligne « Afficher plus de résultats »
  // (loupe) qui déroule le lot complet. On la clique À CHAQUE fois qu'elle apparaît,
  // pour tout dérouler. JAMAIS un contact (le cliquer validerait un destinataire) :
  // on ne clique qu'une ligne SANS adresse dont le texte dit « plus de résultats ».
  var EXPAND_BACKSTOP = 25;                 // garde-fou absolu de clics par recherche
  var expandClicks = 0, lastContacts = -1, noGrowth = 0, expandStalled = false;
  var NARROW_CLICK = '[role="option"],[role="button"],button,[role="menuitem"],a,li,[tabindex]';

  var MORE_RE = /afficher plus|plus de r[eé]sultat|voir plus|show more|more result|rechercher dans l.annuaire|search (the )?directory|weitere ergebnisse|mehr ergebnisse|verzeichnis durchsuchen|mostrar m[aá]s resultados|buscar en el directorio|mostra altri risultati|cerca nella directory|meer resultaten|zoeken in (de )?adreslijst|mostrar mais resultados|pesquisar no diret[oó]rio/i;
  // Un élément qui prend le focus (champ de saisie) : le cliquer ferait quitter le
  // champ « À » et arrêterait le passage. On ne clique jamais ça.
  function takesFocus(el) {
    if (!el || !el.tagName) return false;
    if (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable) return true;
    var r = (el.getAttribute && el.getAttribute("role")) || "";
    return /textbox|searchbox|combobox/.test(r);
  }
  // Le libellé qui porte « plus de résultats » : le texte, sinon aria-label/title
  // (ligne à icône seule). Jamais un contact (adresse), jamais un texte long.
  function moreLabel(el) {
    var tx = (el.textContent || "").trim();
    if (!tx && el.getAttribute) tx = (el.getAttribute("aria-label") || el.getAttribute("title") || "").trim();
    if (!tx || tx.length > 60 || MAIL_RE.test(tx)) return "";
    return MORE_RE.test(tx) ? tx : "";
  }
  // La ligne doit être SOUS le champ « À » (le menu de suggestions y est accroché) :
  // on écarte un « Afficher plus de résultats » d'ailleurs (recherche de courrier...).
  function nearField(el) {
    var f = currentField();
    if (!f || !f.getBoundingClientRect || !el.getBoundingClientRect) return true;
    var a = f.getBoundingClientRect(), b = el.getBoundingClientRect();
    if (!a.width && !a.height && !b.width && !b.height) return true; // pas de mise en page (tests jsdom)
    return b.top >= a.top - 40 && b.top <= a.bottom + 800 && b.right >= a.left - 80 && b.left <= a.right + 80;
  }
  function moreMatch(el) {
    if (!visible(el) || inPanel(el) || takesFocus(el)) return false;
    if (el.closest && el.closest('a[href],[role="main"] [role="document"],[role="article"],article,[aria-label*="Corps du message" i],[aria-label*="Message body" i]')) return false; // jamais un lien ou le texte d'un e-mail
    var own = el.closest && el.closest('[role="listbox"]');
    if (own && isMsgList(own)) return false;          // jamais une ligne de la boîte de réception
    if (!moreLabel(el)) return false;
    return nearField(el);
  }
  // Un VRAI clic (pointeur + souris + click natif) : un simple .click() est souvent
  // ignoré par les boutons Fluent/React d'OWA.
  function realClick(el) {
    var o = { bubbles: true, cancelable: true, composed: true, view: root };
    ["pointerover", "pointerenter", "pointerdown", "mousedown", "pointerup", "mouseup"].forEach(function (type) {
      try { var C = /^pointer/.test(type) && root.PointerEvent ? root.PointerEvent : root.MouseEvent; el.dispatchEvent(new C(type, o)); } catch (e) {}
    });
    try { el.click(); } catch (e) {} // une seule activation (pas de « click » synthétique en plus)
  }
  // Portées du plus proche au plus large : la liste, quelques ancêtres (le callout),
  // puis tout le document en dernier recours (le libellé est très spécifique). La
  // ligne « plus de résultats » est souvent un pied de page FRÈRE de la liste.
  // Jamais « tout le document » : seulement la liste, ses ancêtres proches (le menu
  // de suggestions) et, sans liste, les menus flottants (callout, couche) près du champ.
  var MENUS = '[role="listbox"],[role="dialog"],[role="menu"],[role="presentation"],.ms-Callout,.ms-Layer,[class*="callout" i],[class*="popover" i],[class*="flyout" i],[class*="suggest" i]';
  function moreScopes(box) {
    var scopes = [], s = box || null, hops = 0;
    while (s && hops < 5 && s !== document.body && s !== document.documentElement) { scopes.push(s); s = s.parentElement; hops++; }
    if (!box) [].slice.call(document.querySelectorAll(MENUS)).forEach(function (m) { if (visible(m) && nearField(m) && !inPanel(m)) scopes.push(m); });
    return scopes;
  }
  // La cible RÉELLEMENT cliquable qui porte la phrase : l'élément lui-même s'il est
  // cliquable, sinon un descendant cliquable qui porte aussi la phrase (cas d'un
  // conteneur dont le texte = la phrase mais qui n'a pas de gestionnaire), sinon le
  // plus proche ancêtre cliquable. Évite de cliquer un parent « mort ».
  function clickTarget(node) {
    if (node.matches && node.matches(NARROW_CLICK)) return node;
    var kids = node.querySelectorAll ? [].slice.call(node.querySelectorAll(NARROW_CLICK)) : [];
    for (var i = 0; i < kids.length; i++) if (moreMatch(kids[i])) return kids[i];
    var up = node.closest && node.closest(NARROW_CLICK);
    return (up && !takesFocus(up)) ? up : node;
  }
  // box peut être null (liste pas encore apparue) : on cherche alors la ligne dans
  // tout le document, avec les sélecteurs étroits et la contrainte « sous le champ ».
  function clickMore(box) {
    var scopes = moreScopes(box);
    for (var i = 0; i < scopes.length; i++) {
      var sel = scopes[i] === document ? NARROW_CLICK : NARROW_CLICK + ",div,span"; // document : pas div/span (trop large)
      var nodes = [].slice.call(scopes[i].querySelectorAll(sel));
      for (var j = 0; j < nodes.length; j++) {
        if (!moreMatch(nodes[j])) continue;
        realClick(clickTarget(nodes[j])); return true;
      }
    }
    return false;
  }
  // Combien de contacts (options à adresse) la liste montre en ce moment.
  function contactCount(box) {
    if (!box) return 0;
    return [].slice.call(box.querySelectorAll('[role="option"]')).filter(function (o) { return MAIL_RE.test(o.textContent || ""); }).length;
  }
  function lire() {
    var a = root.AOArmement, field = currentField();
    // Ré-résout la liste à chaque appel (OWA reconstruit les options à chaque frappe).
    var box = resolveListbox(field) || (a && a.isArmed() ? a.dropdown() : null);
    // On clique « plus de résultats » tant qu'elle apparaît — y compris quand la
    // liste n'a encore AUCUN contact : OWA ne propose d'abord que ses contacts
    // récents, et c'est cette ligne qui interroge l'annuaire (le gros du carnet).
    // On ne reclique QU'UNE fois le gain du clic précédent constaté (OWA charge en
    // différé). On arrête si deux lectures d'affilée n'ajoutent plus de contact
    // (anti-boucle sur un « Rechercher dans l'annuaire » permanent), ou au garde-fou.
    if (!expandStalled && expandClicks < EXPAND_BACKSTOP) {
      var now = contactCount(box);
      if (expandClicks > 0) { if (now > lastContacts) { noGrowth = 0; lastContacts = now; } else noGrowth++; }
      else lastContacts = now;
      if (noGrowth >= 2) expandStalled = true;                       // le dernier clic n'ajoute plus rien
      else if (noGrowth === 0 && clickMore(box)) expandClicks++;     // on attend le gain avant de recliquer
    }
    if (!box) return []; // rien de sûr → on ne lit RIEN (jamais la boîte de réception)
    return lireRows(box, SOURCE);
  }

  // --------------------------------------------- ouverture de la rédaction
  function clickByText(sel, words) {
    var all = [].slice.call(document.querySelectorAll(sel));
    for (var i = 0; i < all.length; i++) {
      var el = all[i]; if (!visible(el)) continue;
      var label = (el.getAttribute("aria-label") || el.getAttribute("data-tooltip") || el.textContent || "").trim().toLowerCase();
      for (var w = 0; w < words.length; w++) if (label === words[w] || label.indexOf(words[w]) >= 0) { el.click(); return true; }
    }
    return false;
  }
  function openCompose() {
    var a = root.AOArmement;
    if (a && a.isArmed()) return true; // champ déjà armé par l'utilisateur : rien à ouvrir
    if (findField()) return true; // déjà ouvert
    var cfg = SOURCE === "gmail" ? SEL.gmail : SEL.outlook;
    for (var i = 0; i < (cfg.compose || []).length; i++) { var b = document.querySelector(cfg.compose[i]); if (visible(b)) { b.click(); return true; } }
    return clickByText('[role="button"],button,div[gh],div.T-I,[aria-label]', cfg.composeText || []);
  }

  // -------------------------------------------- présence et pastilles
  function chipScope() {
    var f = currentField();
    if (f && f.closest) { var d = f.closest('[role="dialog"]'); if (d) return d; var p = f; for (var i = 0; i < 5 && p.parentElement; i++) p = p.parentElement; return p; }
    return document;
  }
  // Une pastille destinataire = marqueur générique (data-hovercard-id, role=listitem)
  // ou indice par site, DANS la zone du champ, hors de la liste de suggestions. On
  // ne compte ni les options ni ce qui est dans une [role=listbox].
  function chipSig() {
    var hint = (SOURCE === "gmail" ? SEL.gmail : SEL.outlook).chips || "";
    var sel = '[data-hovercard-id],[role="listitem"]' + (hint ? "," + hint : "");
    var scope = chipScope();
    return [].slice.call(scope.querySelectorAll(sel))
      .filter(function (c) { return c.getAttribute("role") !== "option" && !(c.closest && c.closest('[role="listbox"]')); })
      .map(function (c) { return (c.getAttribute("data-hovercard-id") || c.textContent || "").trim(); })
      .filter(Boolean);
  }
  var baseline = null, listesInitiales = [];
  function newChipAppeared() {
    if (baseline == null) return false;
    var now = chipSig();
    for (var i = 0; i < now.length; i++) if (baseline.indexOf(now[i]) < 0) return true;
    return false;
  }

  // --------------------------------------------------- diagnostic
  var diagCaptures = [];
  // Chaîne d'ancêtres courte (pour situer la ligne « plus de résultats »).
  function ancestry(el) {
    var chain = [], p = el, hops = 0;
    while (p && hops < 6 && p !== document.documentElement) {
      chain.push({ tag: p.tagName, role: p.getAttribute && p.getAttribute("role"), id: p.id || null, cls: (typeof p.className === "string" ? p.className.slice(0, 60) : null) || null });
      p = p.parentElement; hops++;
    }
    return chain;
  }
  // Tout ce qui ressemble à une ligne « Afficher plus de résultats » : où elle est,
  // sa structure, pour viser le clic si l'heuristique la rate sur le vrai OWA.
  function moreCandidates() {
    var out = [], box = resolveListbox(currentField());
    var nodes = [].slice.call(document.querySelectorAll(NARROW_CLICK + ",div,span"));
    for (var i = 0; i < nodes.length && out.length < 8; i++) {
      if (!moreMatch(nodes[i])) continue;
      var e = nodes[i];
      out.push({
        tag: e.tagName, role: e.getAttribute && e.getAttribute("role"),
        id: e.id || null, cls: (typeof e.className === "string" ? e.className.slice(0, 80) : null) || null,
        ariaLabel: e.getAttribute && e.getAttribute("aria-label"), text: (e.textContent || "").trim().slice(0, 60),
        dansLaListe: !!(box && box.contains(e)), dansUneListbox: !!(e.closest && e.closest('[role="listbox"]')),
        ancetres: ancestry(e), outer: (e.outerHTML || "").slice(0, 400)
      });
    }
    return out;
  }
  // Tout élément (quel qu'il soit) dont le libellé dit « plus de résultats » : même
  // non cliquable, même caché — pour voir la vraie structure de la ligne sur ce
  // Outlook et ajuster l'heuristique de clic.
  function moreTexts() {
    var out = [], box = resolveListbox(currentField());
    var nodes = [].slice.call(document.querySelectorAll("*"));
    for (var i = 0; i < nodes.length && out.length < 12; i++) {
      var e = nodes[i];
      if (inPanel(e) || /^(SCRIPT|STYLE|HTML|HEAD|BODY)$/.test(e.tagName)) continue;
      var tx = (e.textContent || "").trim();
      var al = (e.getAttribute && (e.getAttribute("aria-label") || e.getAttribute("title") || e.getAttribute("placeholder"))) || "";
      if (!(tx.length <= 80 && MORE_RE.test(tx)) && !MORE_RE.test(al)) continue;
      var r = {}; try { r = e.getBoundingClientRect(); } catch (x) {}
      out.push({
        tag: e.tagName, role: e.getAttribute("role"), id: e.id || null, cls: (typeof e.className === "string" ? e.className.slice(0, 120) : null) || null,
        ariaLabel: e.getAttribute("aria-label"), title: e.getAttribute("title"), tabindex: e.getAttribute("tabindex"),
        text: tx.slice(0, 80), visible: visible(e), rect: { x: Math.round(r.left || 0), y: Math.round(r.top || 0), w: Math.round(r.width || 0), h: Math.round(r.height || 0) },
        prendLeFocus: takesFocus(e), cliquable: moreMatch(e), dansLaListe: !!(box && box.contains(e)),
        ancetres: ancestry(e), outer: (e.outerHTML || "").slice(0, 700)
      });
    }
    return out;
  }
  // Le conteneur du menu de suggestions (3 niveaux au-dessus de la liste) : la ligne
  // « plus de résultats » y est souvent un pied de page frère de la liste.
  function calloutHtml(box) {
    if (!box) return null;
    var p = box, hops = 0;
    while (p.parentElement && hops < 3 && p.parentElement !== document.body) { p = p.parentElement; hops++; }
    return (p.outerHTML || "").slice(0, 12000);
  }
  function captureField() {
    var el = currentField();
    if (!el) return { field: null, plusDeResultats: moreCandidates(), plusTextes: moreTexts() };
    var box = resolveListbox(el), opts = box ? [].slice.call(box.querySelectorAll('[role="option"]')) : [];
    return {
      field: { tag: el.tagName, contentEditable: el.isContentEditable, type: el.getAttribute("type"), role: el.getAttribute("role"),
        name: (el.closest("[name]") || {}).getAttribute ? el.closest("[name]").getAttribute("name") : null,
        ariaLabel: el.getAttribute("aria-label"), ariaAutocomplete: el.getAttribute("aria-autocomplete"),
        ariaControls: el.getAttribute("aria-controls"), ariaOwns: el.getAttribute("aria-owns"),
        peoplekit: el.getAttribute("peoplekit-id"), value: isField(el) ? el.value : (el.textContent || ""), outer: el.outerHTML.slice(0, 600) },
      listbox: box ? { selectorPortaled: !box.closest('[role="dialog"]'), optionCount: opts.length, contactCount: contactCount(box), outer: box.outerHTML.slice(0, 1200),
        firstOptions: opts.slice(0, 2).map(function (o) { return o.outerHTML.slice(0, 500); }) } : null,
      plusDeResultats: moreCandidates(), // la/les ligne(s) « Afficher plus de résultats » cliquables repérées
      plusTextes: moreTexts(),           // tout ce qui porte le libellé, cliquable ou non
      callout: calloutHtml(box),         // le conteneur du menu, pour voir la vraie structure
      chips: chipSig()
    };
  }
  // Essai de diagnostic : cliquer la ligne « plus de résultats » et dire si on l'a
  // trouvée + combien de contacts avant (le « après » est relu par la capture suivante).
  function plusEssai() {
    var box = resolveListbox(currentField()) || ((root.AOArmement && root.AOArmement.isArmed()) ? root.AOArmement.dropdown() : null);
    var avant = contactCount(box);
    var trouve = clickMore(box);
    return { box: !!box, avant: avant, trouve: trouve };
  }
  function diagnostic(etape) {
    var cap = { etape: etape, source: SOURCE, href: location.href.slice(0, 120) };
    try { Object.assign(cap, captureField()); } catch (e) { cap.error = String(e); }
    diagCaptures.push(cap);
    return cap;
  }
  // Balayage générique : tout ce qui ressemble à un champ destinataire ou à un
  // menu de suggestions, SANS dépendre de mes sélecteurs (justement pour les
  // réparer quand ils ne collent pas à ce vrai Outlook/Gmail).
  function sweep() {
    var out = [], seen = [];
    function add(e, why) {
      if (seen.indexOf(e) >= 0 || out.length >= 150) return; seen.push(e);
      var r = {}; try { r = e.getBoundingClientRect(); } catch (x) {}
      out.push({
        why: why, tag: e.tagName,
        role: e.getAttribute && e.getAttribute("role"),
        type: e.getAttribute && e.getAttribute("type"),
        name: e.getAttribute && e.getAttribute("name"),
        ariaLabel: e.getAttribute && e.getAttribute("aria-label"),
        ariaAutocomplete: e.getAttribute && e.getAttribute("aria-autocomplete"),
        ariaControls: e.getAttribute && e.getAttribute("aria-controls"),
        placeholder: e.getAttribute && e.getAttribute("placeholder"),
        id: e.id || null,
        cls: (typeof e.className === "string" ? e.className.slice(0, 100) : null) || null,
        editable: !!e.isContentEditable,
        visible: !!(r.width && r.height),
        text: (e.textContent || "").trim().slice(0, 50),
        outer: (e.outerHTML || "").slice(0, 350)
      });
    }
    try {
      [].slice.call(document.querySelectorAll("[aria-label],[placeholder],[name]")).forEach(function (e) {
        var s = ((e.getAttribute("aria-label") || "") + " " + (e.getAttribute("placeholder") || "") + " " + (e.getAttribute("name") || "")).toLowerCase();
        if (/destinatair|recipient|\bto\b|\bcc\b|cci|\bà\b|nouveau message|new mail|compose/.test(s)) add(e, "lib:" + s.slice(0, 32));
      });
      ['[role="textbox"]', '[role="combobox"]', "input", "textarea", '[contenteditable="true"]', '[role="listbox"]', '[role="option"]', '[name="to"]'].forEach(function (sel) {
        [].slice.call(document.querySelectorAll(sel)).slice(0, 30).forEach(function (e) { add(e, "sel:" + sel); });
      });
    } catch (e) {}
    return out;
  }
  function diagReport(lu) {
    var data;
    try { data = { source: SOURCE, href: location.href, captures: diagCaptures, lu: lu, balayage: sweep() }; }
    catch (e) { data = { source: SOURCE, captures: diagCaptures, lu: lu, erreur: String(e) }; }
    return "Diagnostic Adresses — " + SOURCE + "\n" + new Date().toString() + "\n\n" +
      JSON.stringify(data, null, 2) + "\n";
  }

  root.AODom = {
    SOURCE: SOURCE,
    taper: function (texte) { expandClicks = 0; lastContacts = -1; noGrowth = 0; expandStalled = false; var el = currentField(); return el ? taper(el, texte) : Promise.resolve({ ok: false, raison: "champ" }); }, // nouvelle recherche → on réautorise « plus de résultats »
    effacer: function (n) { var el = currentField(); return el ? effacer(el, n) : Promise.resolve({ ok: !tape, raison: "champ" }); },
    texteLibre: texteLibre,
    focus: focus,
    lire: lire,
    instantane: function () { baseline = chipSig(); tape = ""; listesInitiales = [].slice.call(document.querySelectorAll('[role="listbox"]')); return true; },
    reste: function () { return tape; }, // texte tapé par Smart Mail resté dans le champ (arrêt sans nettoyage)
    vivant: function () { try { var a = globalThis.chrome || globalThis.browser; return !!(a && a.runtime && a.runtime.id); } catch (e) { return false; } },
    enLigne: function () { return navigator.onLine !== false; },
    nettoyer: function () { return nettoyerSansFocus(); },
    outlookOuvert: function () { return true; },
    activerOutlook: function () { return true; },
    gmailDevant: function () { return !document.hidden; },
    auPremierPlan: function () { return document.hasFocus() && document.visibilityState === "visible"; },
    openCompose: openCompose,
    diagnostic: diagnostic,
    plusEssai: plusEssai,
    diagReport: diagReport,
    // Exposés pour le module d'armement (réutilise la même logique).
    resolveListbox: resolveListbox,
    visible: visible,
    notSearch: notSearch
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
