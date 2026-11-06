// Armement : au lieu de deviner le champ « À » par des sélecteurs (fragile,
// propre à chaque version d'Outlook), on s'accroche au champ que L'UTILISATEUR
// vient d'utiliser. Dès qu'il tape dans un champ éditable ET qu'une liste de
// suggestions CONTENANT DES ADRESSES apparaît, on « arme » ce champ-là et sa
// liste. Le reste (adapter-dom) pointe alors dessus. Marche sur n'importe quelle
// version, parce que ça repose sur le vrai champ et la vraie liste, pas sur un
// nom de classe. Exposé via window.AOArmement.

(function (root) {
  "use strict";

  var AO_WAIT = (root.__AO_TEST && root.__AO_TEST.waitMs != null) ? root.__AO_TEST.waitMs : null;
  function MAIL() { return (root.AOLecture && root.AOLecture.MAIL) || /[^\s@]+@[^\s@]+\.[A-Za-z]{2,}/; }
  function dom() { return root.AODom || {}; }

  var lastEditable = null;   // dernier champ éditable focalisé par l'utilisateur
  var typedBuffer = "";
  var armedField = null, armedDropdown = null;
  var suspended = false;
  var armedCbs = [], disarmedCbs = [];
  var scanTimer = null;

  function editable(el) {
    return !!el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable);
  }
  function visible(el) {
    var d = dom();
    if (d.visible) return d.visible(el);
    if (!el) return false;
    if (el.getAttribute && el.getAttribute("aria-hidden") === "true") return false;
    return el.offsetParent !== null || (el.getClientRects && el.getClientRects().length > 0);
  }
  function notSearch(el) {
    var d = dom();
    if (d.notSearch) return d.notSearch(el);
    if (!el) return false;
    if (el.closest && el.closest('[role="search"],[role="banner"],header')) return false;
    var al = ((el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("placeholder"))) || "").toLowerCase();
    return !/search|recherch/.test(al);
  }
  function fieldText(el) {
    if (!el) return "";
    return (el.value != null ? String(el.value) : (el.textContent || "")).trim();
  }
  function hasMail(box) { return !!box && MAIL().test(box.textContent || ""); }

  // La liste liée au champ : d'abord l'ARIA (AODom.resolveListbox), sinon une
  // [role=listbox] visible à adresses, sinon le parent d'une option à adresse.
  function resolve(el) {
    var d = dom();
    // resolveListbox (adapter-dom) écarte déjà la boîte de réception et notre
    // panneau : on s'appuie dessus plutôt que de rescanner toutes les listbox.
    var box = d.resolveListbox ? d.resolveListbox(el) : null;
    return (box && visible(box) && hasMail(box)) ? box : null;
  }

  // La liste « répond » à ce que l'utilisateur a tapé ? (un mot d'option commence
  // par la saisie) — pour écarter un menu périmé ou sans rapport.
  function respondsTo(box, q) {
    if (!q) return true;
    var ql = q.toLowerCase();
    var opts = box.querySelectorAll('[role="option"]');
    var pool = opts.length ? [].slice.call(opts) : [box];
    for (var i = 0; i < pool.length; i++) {
      var words = (pool[i].textContent || "").toLowerCase().split(/[^a-z0-9]+/);
      for (var w = 0; w < words.length; w++) if (words[w] && words[w].indexOf(ql) === 0) return true;
    }
    return false;
  }

  function fire(cbs, arg) { for (var i = 0; i < cbs.length; i++) { try { cbs[i](arg); } catch (e) {} } }

  function tryArm() {
    if (suspended || armedField) return;
    var el = lastEditable;
    if (!editable(el) || !document.contains(el) || !visible(el) || !notSearch(el)) return;
    var box = resolve(el);
    if (!box || !hasMail(box)) return;
    if (box.closest && box.closest('[data-testid="ao-panel"]')) return; // pas notre propre panneau
    if (!respondsTo(box, typedBuffer)) return;
    armedField = el; armedDropdown = box;
    fire(armedCbs, { field: el, dropdown: box });
  }

  // L'autocomplétion arrive en différé : on sonde quelques fois après une frappe.
  function scheduleScan() {
    if (armedField || scanTimer) return;
    var tries = 0, step = AO_WAIT != null ? AO_WAIT : 140;
    scanTimer = setInterval(function () {
      tries++; tryArm();
      if (armedField || tries >= 25) { clearInterval(scanTimer); scanTimer = null; } // ~3,7 s : certaines listes OWA tardent
    }, step);
  }

  // ---- Démarrage automatique : si le curseur reste dans un champ de type
  // destinataire SANS que l'utilisateur tape (nouveau message → curseur déjà dans
  // « À »), on arme ce champ au bout de ~2 s. Le moteur tapera lui-même la lettre.
  var autoTimer = null;
  function recipientish(el) {
    if (!editable(el) || !notSearch(el)) return false;
    if (el.closest && el.closest('[data-testid="ao-panel"]')) return false; // pas notre propre panneau
    var al = ((el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("placeholder") || el.getAttribute("name"))) || "").trim().toLowerCase();
    if (/corps|body|\bobjet\b|subject|message|signature/.test(al)) return false; // ni le corps, ni l'objet
    // Gmail : champ destinataire reconnaissable par son conteneur « to/cc/bcc » ou
    // par l'attribut peoplekit, même si le libellé est inhabituel.
    if (el.getAttribute && el.getAttribute("peoplekit-id")) return true;
    if (el.closest && el.closest('div[name="to"],div[name="cc"],div[name="bcc"]')) return true;
    // Libellés destinataire courants. Le « À » français fait UNE lettre, que \b ne
    // borne pas (« à » n'est pas un caractère de mot ASCII) : on le teste à part.
    if (al === "à" || al === "a" || al === "to" || al === "cc" || al === "cci" || al === "bcc") return true;
    if (/destinatair|recipient|\bto\b|\bcc\b|cci|\bbcc\b/.test(al)) return true;
    if (/(^|[^a-z])à([^a-z]|$)/.test(al)) return true; // « à : », « champ à », etc.
    if (el.getAttribute && el.getAttribute("aria-autocomplete")) return true;
    if (el.getAttribute && /combobox/.test(el.getAttribute("role") || "")) return true;
    var inDlg = el.closest && el.closest('[role="dialog"]');
    if (inDlg && (el.tagName === "INPUT" || el.tagName === "TEXTAREA")) return true;
    if (inDlg && el.isContentEditable) return true; // contenteditable de rédaction (hors corps/objet)
    return false;
  }
  function forceArm(el) {
    if (suspended || armedField) return;
    if (!editable(el) || !document.contains(el) || !visible(el) || !notSearch(el)) return;
    armedField = el; armedDropdown = null; // la liste apparaîtra quand le moteur tapera
    fire(armedCbs, { field: el, dropdown: null });
  }

  // ---- Démarrage MANUEL (« Démarrer ici ») : on arme SANS condition le champ où
  // est (ou était à l'instant) le curseur. Pas de liste, pas de « est-ce le bon
  // champ » : le clic de l'utilisateur fait foi. Seule exclusion gardée : la barre
  // de recherche et notre propre panneau (on ne tape jamais là-dedans). N'émet PAS
  // de rappel : c'est le bouton qui démarre, sans compte à rebours.
  function pickHere() {
    var el = document.activeElement;
    if (editable(el) && visible(el) && notSearch(el) && !(el.closest && el.closest('[data-testid="ao-panel"]'))) return el;
    if (editable(lastEditable) && document.contains(lastEditable) && visible(lastEditable) && notSearch(lastEditable)) return lastEditable;
    return null;
  }
  function armHere() {
    cancelAuto(); // un démarrage manuel (« Démarrer ici ») annule un démarrage auto en attente
    if (armedField && document.contains(armedField)) return true;
    var el = pickHere();
    if (!el) return false;
    armedField = el; armedDropdown = null;
    return true;
  }
  function scheduleAuto(el) {
    if (autoTimer) { clearTimeout(autoTimer); autoTimer = null; }
    if (!recipientish(el)) return;
    autoTimer = setTimeout(function () {
      autoTimer = null;
      if (suspended || armedField || document.activeElement !== el || !document.contains(el) || fieldText(el)) return;
      forceArm(el);
    }, root.__AO_TEST ? 400 : 2000); // délai propre (pas AO_WAIT) pour que la frappe manuelle prime
  }
  function cancelAuto() { if (autoTimer) { clearTimeout(autoTimer); autoTimer = null; } }

  document.addEventListener("focusin", function (e) {
    if (suspended) return;
    if (editable(e.target) && notSearch(e.target)) { lastEditable = e.target; typedBuffer = fieldText(e.target); scheduleAuto(e.target); }
    else cancelAuto();
  }, true);
  document.addEventListener("focusout", cancelAuto, true);

  function onUserType(e) {
    if (suspended || !e.isTrusted) return; // ignore la frappe SYNTHÉTIQUE du moteur
    cancelAuto(); // l'utilisateur tape : la voie normale (liste réelle) prend le relais
    if (editable(e.target)) { lastEditable = e.target; typedBuffer = fieldText(e.target); }
    else if (lastEditable) typedBuffer = fieldText(lastEditable);
    scheduleScan();
  }
  document.addEventListener("keydown", onUserType, true);
  document.addEventListener("input", onUserType, true);

  function checkAlive() {
    if (armedField && (!document.contains(armedField) || !visible(armedField))) {
      armedField = null; armedDropdown = null; fire(disarmedCbs, {});
    }
  }

  root.AOArmement = {
    isArmed: function () { checkAlive(); return !!armedField; },
    field: function () { return armedField; },
    dropdown: function () { return (armedDropdown && document.contains(armedDropdown)) ? armedDropdown : null; },
    typed: function () { return typedBuffer; },
    armHere: armHere,          // armement manuel inconditionnel (bouton « Démarrer ici »)
    suspend: function () { suspended = true; },
    resume: function () { suspended = false; },
    disarm: function () { armedField = null; armedDropdown = null; },
    onArmed: function (cb) { armedCbs.push(cb); },
    onDisarmed: function (cb) { disarmedCbs.push(cb); }
  };

  // Hook de test : arme le champ actuellement focalisé (déterministe en e2e).
  document.addEventListener("ao-arm", function () {
    lastEditable = document.activeElement;
    typedBuffer = fieldText(lastEditable);
    tryArm();
  });
})(typeof globalThis !== "undefined" ? globalThis : this);
