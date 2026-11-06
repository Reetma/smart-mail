// Gestes dans Outlook web, comme une personne qui fait des copier-coller :
// cliquer « Nouveau message », coller l'adresse dans « À » (validée par « ; »)
// ou un paquet d'adresses dans « Cci », coller l'objet, coller le message avec
// ses liens, au besoin retirer le texte de la signature et cocher « Importance
// haute », cliquer « Envoyer », attendre que la rédaction se referme. Aucun appel
// réseau : seule la page Outlook ouverte.
//
// Exposé dans window.PP.Outlook (lu par moteur.js et panneau.js).

(function (root) {
  "use strict";
  var PP = (root.PP = root.PP || {});

  // Pages de test : <html data-pp-test="rapide"> divise toutes les attentes par 100.
  function facteur() {
    try { return document.documentElement.dataset.ppTest === "rapide" ? 0.01 : 1; } catch (e) { return 1; }
  }
  // Attente qui tient aussi onglet caché (Outlook derrière une autre fenêtre ou un
  // autre onglet) : Chrome y ralentit les minuteries de la page, jusqu'à un réveil
  // par minute. Le service de l'extension (sw.js) n'est pas ralenti : chaque tranche
  // (20 s au plus, pour qu'il reste éveillé) finit au premier des deux réveils.
  var api = root.chrome || root.browser;
  function tranche(ms) {
    return new Promise(function (r) {
      setTimeout(r, ms);
      try { api.runtime.sendMessage({ pp: "reveil", ms: ms }).then(r, r); } catch (e) {}
    });
  }
  async function attendre(ms) {
    var fin = Date.now() + Math.max(0, ms * facteur());
    for (var reste = fin - Date.now(); reste > 0; reste = fin - Date.now()) await tranche(Math.min(reste, 20000));
  }
  function hasard(min, max) { return min + Math.random() * (max - min); }

  var SEL = {
    nouveau: ['[aria-label="Nouveau message"]', '[aria-label="New mail"]', '[aria-label="New message"]', '[aria-label="Nouveau courrier"]', '[aria-label="Nouveau courriel"]', '[aria-label="New email"]', '[aria-label="Neue E-Mail"]', '[aria-label="Neue Nachricht"]', '[aria-label="Correo nuevo"]', '[aria-label="Mensaje nuevo"]', '[aria-label="Nuova e-mail"]', '[aria-label="Nuovo messaggio"]', '[aria-label="Nieuwe e-mail"]', '[aria-label="Nieuw bericht"]', '[aria-label="Novo email"]', '[aria-label="Nova mensagem"]', 'button[title="Nouveau message"]', 'button[title="New mail"]', 'button[title="New message"]', 'button[title="Nouveau courrier"]', 'button[title="Nouveau courriel"]', 'button[title="New email"]', 'button[title="Neue E-Mail"]', 'button[title="Neue Nachricht"]', 'button[title="Correo nuevo"]', 'button[title="Mensaje nuevo"]', 'button[title="Nuova e-mail"]', 'button[title="Nuovo messaggio"]', 'button[title="Nieuwe e-mail"]', 'button[title="Nieuw bericht"]', 'button[title="Novo email"]', 'button[title="Nova mensagem"]'],
    nouveauTexte: ["nouveau message", "new mail", "new message", "nouveau courrier", "nouveau courriel", "new email", "neue e-mail", "neue nachricht", "correo nuevo", "mensaje nuevo", "nuova e-mail", "nuovo messaggio", "nieuwe e-mail", "nieuw bericht", "novo email", "nova mensagem"],
    a: ['div[contenteditable="true"][aria-label="À"]', 'div[contenteditable="true"][aria-label="To"]', 'div[contenteditable="true"][aria-label="An"]', 'div[contenteditable="true"][aria-label="Para"]', 'div[contenteditable="true"][aria-label="A"]', 'div[contenteditable="true"][aria-label="Aan"]', '[role="textbox"][aria-label="À"]', '[role="textbox"][aria-label="To"]', '[role="textbox"][aria-label="An"]', '[role="textbox"][aria-label="Para"]', '[role="textbox"][aria-label="A"]', '[role="textbox"][aria-label="Aan"]', 'input[aria-label="À"]', 'input[aria-label="To"]', 'input[aria-label="An"]', 'input[aria-label="Para"]', 'input[aria-label="A"]', 'input[aria-label="Aan"]', '[aria-label="À"] input', '[aria-label="To"] input', '[aria-label="An"] input', '[aria-label="Para"] input', '[aria-label="A"] input', '[aria-label="Aan"] input', 'input[aria-label="Destinataires"]', '.ms-BasePicker-input'],
    objet: ['input[aria-label="Ajouter un objet"]', 'input[aria-label="Add a subject"]', 'input[aria-label="Objet"]', 'input[aria-label="Subject"]', 'input[aria-label="Betreff"]', 'input[aria-label="Betreff hinzufügen"]', 'input[aria-label="Asunto"]', 'input[aria-label="Agregar un asunto"]', 'input[aria-label="Agregue un asunto"]', 'input[aria-label="Oggetto"]', 'input[aria-label="Aggiungi un oggetto"]', 'input[aria-label="Onderwerp"]', 'input[aria-label="Onderwerp toevoegen"]', 'input[aria-label="Assunto"]', 'input[aria-label="Adicionar um assunto"]', 'input[placeholder="Ajouter un objet"]', 'input[placeholder="Add a subject"]', 'input[placeholder="Objet"]', 'input[placeholder="Subject"]', 'input[placeholder="Betreff"]', 'input[placeholder="Betreff hinzufügen"]', 'input[placeholder="Asunto"]', 'input[placeholder="Agregar un asunto"]', 'input[placeholder="Agregue un asunto"]', 'input[placeholder="Oggetto"]', 'input[placeholder="Aggiungi un oggetto"]', 'input[placeholder="Onderwerp"]', 'input[placeholder="Onderwerp toevoegen"]', 'input[placeholder="Assunto"]', 'input[placeholder="Adicionar um assunto"]', 'input[aria-label*="objet" i]', 'input[aria-label*="subject" i]', 'input[placeholder*="objet" i]', 'input[placeholder*="subject" i]'],
    corps: ['div[contenteditable="true"][aria-label="Corps du message"]', 'div[contenteditable="true"][aria-label="Message body"]', 'div[contenteditable="true"][aria-label="Nachrichtentext"]', 'div[contenteditable="true"][aria-label="Cuerpo del mensaje"]', 'div[contenteditable="true"][aria-label="Corpo del messaggio"]', 'div[contenteditable="true"][aria-label="Berichttekst"]', 'div[contenteditable="true"][aria-label="Corpo da mensagem"]', 'div[contenteditable="true"][aria-label*="corps" i]', 'div[contenteditable="true"][aria-label*="body" i]', 'div[role="textbox"][contenteditable="true"][aria-multiline="true"]'],
    envoyer: ['button[aria-label^="Envoyer ("]', 'button[title^="Envoyer ("]', 'button[aria-label^="Send ("]', 'button[title^="Send ("]', 'button[aria-label^="Senden ("]', 'button[title^="Senden ("]', 'button[aria-label^="Enviar ("]', 'button[title^="Enviar ("]', 'button[aria-label^="Invia ("]', 'button[title^="Invia ("]', 'button[aria-label^="Verzenden ("]', 'button[title^="Verzenden ("]', 'button[aria-label="Envoyer"]', 'button[aria-label="Send"]', 'button[aria-label="Senden"]', 'button[aria-label="Enviar"]', 'button[aria-label="Invia"]', 'button[aria-label="Verzenden"]', '[role="button"][aria-label="Envoyer"]', '[role="button"][aria-label="Send"]', '[role="button"][aria-label="Senden"]', '[role="button"][aria-label="Enviar"]', '[role="button"][aria-label="Invia"]', '[role="button"][aria-label="Verzenden"]', 'button[title^="Envoyer"]', 'button[title^="Send"]', 'button[title^="Senden"]', 'button[title^="Enviar"]', 'button[title^="Invia"]', 'button[title^="Verzenden"]'],
    envoyerTexte: ["envoyer", "send", "senden", "enviar", "invia", "verzenden"],
    ignorer: ['button[aria-label="Abandonner"]', 'button[aria-label="Ignorer"]', 'button[aria-label="Discard"]', 'button[aria-label="Verwerfen"]', 'button[aria-label="Descartar"]', 'button[aria-label="Rimuovi"]', 'button[aria-label="Negeren"]', 'button[title^="Abandonner"]', 'button[title^="Ignorer"]', 'button[title^="Discard"]', 'button[title^="Verwerfen"]', 'button[title^="Descartar"]', 'button[title^="Rimuovi"]', 'button[title^="Negeren"]'],
    cci: ['div[contenteditable="true"][aria-label="Cci"]', 'div[contenteditable="true"][aria-label="Bcc"]', 'div[contenteditable="true"][aria-label="BCC"]', 'div[contenteditable="true"][aria-label="CCO"]', 'div[contenteditable="true"][aria-label="Cco"]', 'div[contenteditable="true"][aria-label="Ccn"]', '[role="textbox"][aria-label="Cci"]', '[role="textbox"][aria-label="Bcc"]', '[role="textbox"][aria-label="BCC"]', '[role="textbox"][aria-label="CCO"]', '[role="textbox"][aria-label="Cco"]', '[role="textbox"][aria-label="Ccn"]', 'input[aria-label="Cci"]', 'input[aria-label="Bcc"]', 'input[aria-label="BCC"]', 'input[aria-label="CCO"]', 'input[aria-label="Cco"]', 'input[aria-label="Ccn"]', '[aria-label="Cci"] input', '[aria-label="Bcc"] input', '[aria-label="BCC"] input', '[aria-label="CCO"] input', '[aria-label="Cco"] input', '[aria-label="Ccn"] input'],
    cc: ['div[contenteditable="true"][aria-label="Cc"]', 'div[contenteditable="true"][aria-label="CC"]', '[role="textbox"][aria-label="Cc"]', '[role="textbox"][aria-label="CC"]', 'input[aria-label="Cc"]', 'input[aria-label="CC"]', '[aria-label="Cc"] input', '[aria-label="CC"] input'],
  };
  // Libellés (aria-label, info-bulle ou texte) des commandes cherchées.
  var LIB = {
    montrerCci: /^(cci|bcc|cco|ccn)$|\b(cci|bcc|cco|ccn)\b/i, // « Cci », « Afficher la ligne Cci », « Show Bcc line », « Mostrar la línea CCO »…
    importanceHaute: /importance\s*:?\s*(haute|élevée|elevee|high)|(haute|high)\s+importance|hohe wichtigkeit|wichtigkeit:?\s*hoch|importancia alta|priorit[àa] alta|urgentie:?\s*hoog|alta prioridade/i,
    importance: /^(définir l'importance|definir l'importance|set importance|importance|wichtigkeit festlegen|establecer importancia|imposta priorit[àa]|urgentie instellen|definir prioridade)$/i,
    choixHaute: /^(haute|élevée|elevee|high|importance haute|importance élevée|high importance|hoch|hohe wichtigkeit|alta|importancia alta|priorit[àa] alta|hoog|urgentie: hoog|alta prioridade)$/i,
    plusOptions: /^(plus d.options|more options|autres options|options supplémentaires|plus de commandes|more commands|weitere optionen|más opciones|altre opzioni|meer opties|mais opções|\.\.\.|…)$/i,
    ongletMessage: /^(message|nachricht|mensaje|messaggio|bericht|mensagem)$/i,
    ongletOptions: /^(options|optionen|opciones|opzioni|opties|opções)$/i,
  };

  var MAIL = /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/;

  // ------------------------------------------------------------ repérage
  // Visible à l'écran. On ne se fie pas à aria-hidden : quand une fenêtre Outlook
  // s'ouvre par-dessus, tout le reste est marqué aria-hidden alors qu'il est affiché.
  function visible(el) {
    if (!el || !el.isConnected) return false;
    if (el.closest && el.closest("[data-pp-panneau]")) return false; // jamais notre propre panneau
    if (!el.getClientRects || !el.getClientRects().length) return false;
    var r = el.getBoundingClientRect();
    if (!r.width && !r.height) return false;
    try { var cs = root.getComputedStyle(el); if (cs.visibility === "hidden" || cs.display === "none") return false; } catch (e) {}
    return true;
  }
  function horsRecherche(el) {
    if (el.closest && el.closest('[role="search"],[role="banner"],header')) return false;
    var al = ((el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("placeholder"))) || "").toLowerCase();
    return !/search|recherch/.test(al);
  }
  function premier(selecteurs, portee) {
    portee = portee || document;
    for (var i = 0; i < selecteurs.length; i++) {
      var liste = portee.querySelectorAll(selecteurs[i]);
      for (var j = 0; j < liste.length; j++) if (visible(liste[j]) && horsRecherche(liste[j])) return liste[j];
    }
    return null;
  }
  function parTexte(mots, portee) {
    var tous = (portee || document).querySelectorAll('button,[role="button"],[role="menuitem"]');
    for (var i = 0; i < tous.length; i++) {
      var el = tous[i];
      if (!visible(el)) continue;
      var lib = (el.getAttribute("aria-label") || el.getAttribute("title") || el.textContent || "").trim().toLowerCase();
      for (var m = 0; m < mots.length; m++) if (lib === mots[m]) return el;
    }
    return null;
  }

  // Fenêtre Outlook ouverte par-dessus (modale) qui ne contient pas la rédaction :
  // tant qu'elle est là, cliquer ou taper ailleurs ne ferait rien de bon.
  // Réellement à l'écran : dans la fenêtre, d'une taille raisonnable, pas transparent.
  // (Les compléments comme Letsignit hébergent parfois leur exécution dans une
  // fenêtre cachée ou hors de l'écran : elle ne gêne pas.)
  function aLEcran(el) {
    var r = el.getBoundingClientRect();
    if (r.width < 40 || r.height < 30) return false;
    if (r.right <= 0 || r.bottom <= 0 || r.left >= root.innerWidth || r.top >= root.innerHeight) return false;
    try { if (parseFloat(root.getComputedStyle(el).opacity) === 0) return false; } catch (e) {}
    return true;
  }
  function dialogueBloquant() {
    var ds = document.querySelectorAll('[role="dialog"][aria-modal="true"],[role="alertdialog"],.ms-Modal.is-open');
    for (var i = 0; i < ds.length; i++) {
      var d = ds[i];
      if (!visible(d) || !aLEcran(d)) continue;
      if (premier(SEL.a, d) && premier(SEL.objet, d) && premier(SEL.corps, d)) continue; // c'est la rédaction elle-même (fenêtre séparée)
      return d;
    }
    return null;
  }
  function decrireDialogue(d) {
    var boutons = [].slice.call(d.querySelectorAll('button,[role="button"]')).filter(visible)
      .map(function (b) { return (b.getAttribute("aria-label") || b.textContent || "").trim().replace(/\s+/g, " ").slice(0, 40); }).filter(Boolean);
    return { texte: (d.textContent || "").trim().replace(/\s+/g, " ").slice(0, 300), boutons: boutons.slice(0, 8) };
  }
  function nomDialogue(d) { var t = decrireDialogue(d).texte; return t ? " « " + t.slice(0, 80) + " »" : ""; }
  // Une fenêtre par-dessus (souvent passagère : un complément comme Letsignit qui
  // pose la signature) : on attend qu'elle se referme. null si la voie est libre,
  // sinon le message d'arrêt (fenêtre toujours là après maxMs).
  async function attendreFenetre(maxMs) {
    var d = dialogueBloquant();
    if (!d) return null;
    if (PP.Outlook.surAttente) PP.Outlook.surAttente("Attente : Outlook affiche" + nomDialogue(d) + "…");
    var fin = Date.now() + maxMs * facteur();
    while (Date.now() < fin) {
      await attendre(250);
      d = dialogueBloquant();
      if (!d) { await attendre(hasard(400, 900)); return null; }
    }
    var info = decrireDialogue(d);
    return "Outlook affiche une fenêtre par-dessus depuis " + Math.round(maxMs / 1000) + " s" + (info.texte ? " : « " + info.texte.slice(0, 160) + " »" : "") + ". Fermez-la, puis relancez.";
  }

  // RÉDACTION ANCRÉE : celle que Smart Mail vient d'ouvrir lui-même (repérée une
  // fois pour toutes dans ouvrirRedaction). Pendant un envoi, TOUT se cherche
  // dedans et nulle part ailleurs : jamais le brouillon de l'utilisateur, une
  // réponse en ligne, une autre rédaction ou un bouton « Envoyer » d'ailleurs.
  var redaction = null;
  // Ce que Smart Mail a mis dans SA rédaction : revérifié juste avant « Envoyer ».
  var prevu = {};
  var dernierMessage = ""; // texte (compact) du dernier message envoyé
  function ancree() { return !!(redaction && redaction.isConnected && visible(redaction)); }
  function dansRedaction(sel) { return ancree() ? premier(sel, redaction) : null; }
  function champA() { return redaction ? dansRedaction(SEL.a) : premier(SEL.a); }
  function champCci() { return redaction ? dansRedaction(SEL.cci) : premier(SEL.cci); }
  function champCc() { return redaction ? dansRedaction(SEL.cc) : premier(SEL.cc); }
  function champObjet() { return redaction ? dansRedaction(SEL.objet) : premier(SEL.objet); }
  function champCorps() { return redaction ? dansRedaction(SEL.corps) : premier(SEL.corps); }
  // Une rédaction est-elle ouverte quelque part (détection, pas d'action) ?
  function redactionOuverte() { return !!(premier(SEL.objet) || premier(SEL.a)); }
  // Notre rédaction est-elle encore là, entière ?
  function redactionIntacte() { return ancree() && !!(premier(SEL.a, redaction) && premier(SEL.objet, redaction) && premier(SEL.corps, redaction)); }
  var PERDUE = "La rédaction ouverte par Smart Mail n'est plus là (fermée, remplacée ou masquée) : rien n'est envoyé.";
  // L'utilisateur est-il en train d'écrire AILLEURS dans Outlook (onglet visible et
  // actif, curseur dans une zone de saisie hors de notre rédaction) ? Dans ce cas on
  // ne lui prend pas le curseur : ses frappes partiraient dans notre mail.
  function editable(el) { return !!el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable); }
  function utilisateurOccupe() {
    var a = document.activeElement;
    if (document.hidden || !document.hasFocus() || !editable(a)) return false;
    if (a.closest && a.closest("[data-pp-panneau],[data-testid=\"ao-panel\"]")) return false;
    return !(redaction && redaction.contains(a));
  }
  var OCCUPE = "Envoi mis en pause : vous écrivez dans Outlook depuis 30 minutes. Rien n'est envoyé ; cliquez « Reprendre » quand vous voulez.";
  // On attend (sans prendre le curseur) que l'utilisateur clique ailleurs ou change
  // d'onglet / d'application ; au-delà de 30 minutes, pause.
  async function attendreLibre() {
    if (!utilisateurOccupe()) return null;
    if (PP.Outlook.surAttente) PP.Outlook.surAttente("En attente : vous écrivez dans Outlook. L'envoi reprend dès que vous cliquez ailleurs ou changez d'onglet.");
    var libre = await attendreQue(function () { return !utilisateurOccupe(); }, 30 * 60000, 400);
    return libre ? null : OCCUPE;
  }
  // Une VRAIE frappe au clavier (ou un collage, un glisser-déposer de l'utilisateur)
  // qui atteint notre rédaction : le message n'est plus celui prévu, rien ne part.
  // (Les gestes de Smart Mail sont synthétiques : isTrusted = false.)
  ["keydown", "paste", "drop", "cut"].forEach(function (type) {
    document.addEventListener(type, function (e) {
      if (!e.isTrusted || !redaction || !e.target || !redaction.contains(e.target)) return;
      if (type === "keydown" && /^(Shift|Control|Alt|Meta|CapsLock|Tab|Escape|ArrowUp|ArrowDown|ArrowLeft|ArrowRight|PageUp|PageDown|Home|End)$/.test(e.key)) return;
      prevu.souille = true;
    }, true);
  });

  // Zone de la rédaction : la rédaction ancrée ; à défaut (diagnostic), le plus
  // petit ancêtre commun qui contient l'objet et le bouton Envoyer.
  function zoneRedaction() {
    if (redaction) return ancree() ? redaction : null;
    var o = premier(SEL.objet) || premier(SEL.a);
    if (!o) return null;
    var p = o;
    for (var i = 0; i < 25 && p; i++) {
      if (premier(SEL.envoyer, p) || parTexte(SEL.envoyerTexte, p)) return p;
      p = p.parentElement;
    }
    return null;
  }
  // Tous les éléments (visibles ou non) qui correspondent : pour reconnaître ceux
  // qui existaient AVANT l'ouverture de notre rédaction.
  function tous(selecteurs) {
    var out = [];
    for (var i = 0; i < selecteurs.length; i++) [].push.apply(out, [].slice.call(document.querySelectorAll(selecteurs[i])));
    return out;
  }
  function premierNouveau(selecteurs, anciens) {
    for (var i = 0; i < selecteurs.length; i++) {
      var liste = document.querySelectorAll(selecteurs[i]);
      for (var j = 0; j < liste.length; j++) if (anciens.indexOf(liste[j]) < 0 && visible(liste[j]) && horsRecherche(liste[j])) return liste[j];
    }
    return null;
  }
  // Le conteneur de la nouvelle rédaction : le plus petit ancêtre de son « À » qui
  // contient aussi SON objet, SON corps et un bouton « Envoyer ».
  function conteneurDe(a, objet, corps) {
    for (var p = a, i = 0; p && p !== document.body && i < 40; p = p.parentElement, i++) {
      if (!(p.contains(objet) && p.contains(corps))) continue;
      if (saisieEtrangere(p, [a, objet, corps])) return null; // on déborderait sur une autre zone (Teams, recherche…)
      if (premier(SEL.envoyer, p) || parTexte(SEL.envoyerTexte, p)) return p;
    }
    return null;
  }
  // Une zone de saisie visible dans p qui n'appartient pas à la rédaction (ni À,
  // Cc, Cci, Objet, Corps, ni leurs lignes) : discussion Teams, recherche, autre message.
  function saisieEtrangere(p, propres) {
    // La « ligne » de chaque champ du message : jusqu'à 4 parents (le vrai Outlook
    // entoure « Cci » d'un span.fui-Input qui contient aussi un petit input vide).
    var lignes = [];
    propres.concat(tous(SEL.cc), tous(SEL.cci)).filter(function (x) { return x && p.contains(x); }).forEach(function (x) {
      for (var q = x.parentElement, k = 0; q && q !== p && k < 4; q = q.parentElement, k++) lignes.push(q);
    });
    var champs = p.querySelectorAll('input:not([type="hidden"]):not([type="checkbox"]):not([type="radio"]):not([type="button"]):not([type="submit"]),textarea,[contenteditable="true"],[role="textbox"],[role="searchbox"],[role="combobox"]');
    for (var i = 0; i < champs.length; i++) {
      var c = champs[i];
      if (!visible(c)) continue;
      if (propres.some(function (x) { return x === c || x.contains(c) || c.contains(x); })) continue;
      if (lignes.some(function (l) { return l.contains(c); })) continue;
      if (correspond(c, SEL.cc) || correspond(c, SEL.cci) || correspond(c, SEL.a)) continue;
      return c;
    }
    return null;
  }

  // ------------------------------------------------------------ souris
  function survoler(el) {
    var o = { bubbles: true, cancelable: true, composed: true, view: root };
    ["pointerover", "pointerenter", "mouseover", "mouseenter", "pointermove", "mousemove"].forEach(function (type) {
      try { var C = /^pointer/.test(type) && root.PointerEvent ? root.PointerEvent : root.MouseEvent; el.dispatchEvent(new C(type, o)); } catch (e) {}
    });
  }
  // Un vrai clic (pointeur + souris + activation) : un simple .click() est parfois ignoré par Outlook.
  // Rend true (clic complet), « absent » (l'élément avait disparu AVANT l'appui :
  // aucun geste) ou « incertain » (disparu pendant l'appui : issue inconnue).
  async function cliquer(el) {
    try { el.scrollIntoView({ block: "nearest" }); } catch (e) {}
    survoler(el);
    await attendre(hasard(120, 380));
    if (!el.isConnected) return "absent";
    var o = { bubbles: true, cancelable: true, composed: true, view: root, button: 0, buttons: 1 };
    ["pointerdown", "mousedown"].forEach(function (type) {
      try { var C = /^pointer/.test(type) && root.PointerEvent ? root.PointerEvent : root.MouseEvent; el.dispatchEvent(new C(type, o)); } catch (e) {}
    });
    if (el.focus) { try { el.focus(); } catch (e) {} }
    await attendre(hasard(40, 120));
    o.buttons = 0;
    if (!el.isConnected) return "incertain"; // disparu entre l'appui et le relâché
    ["pointerup", "mouseup"].forEach(function (type) {
      try { var C = /^pointer/.test(type) && root.PointerEvent ? root.PointerEvent : root.MouseEvent; el.dispatchEvent(new C(type, o)); } catch (e) {}
    });
    if (!el.isConnected) return "incertain";
    try { el.click(); } catch (e) {}
    return true;
  }

  // ------------------------------------------------------------ saisie
  function estChamp(el) { return el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA"); }
  function valeurNative(el, v) {
    var proto = el.tagName === "TEXTAREA" ? root.HTMLTextAreaElement.prototype : root.HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value").set.call(el, v);
  }
  function evenementSaisie(el, type, data) {
    el.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true, inputType: type, data: data }));
  }
  function curseurDedans(el) {
    if (estChamp(el)) return;
    var s = root.getSelection();
    if (s && s.rangeCount && el.contains(s.anchorNode)) return;
    var r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    s.removeAllRanges();
    s.addRange(r);
  }
  function focusSur(el) { var a = document.activeElement; return !!a && (a === el || (el.contains && el.contains(a))); }
  function insererTexte(el, texte) {
    if (!focusSur(el)) return false; // une fenêtre a pris le curseur : on ne tape pas chez elle
    curseurDedans(el);
    var ok;
    try { ok = document.execCommand("insertText", false, texte); } catch (e) { ok = false; }
    if (ok) return true;
    if (estChamp(el)) {
      var d = el.selectionStart == null ? el.value.length : el.selectionStart;
      var f = el.selectionEnd == null ? el.value.length : el.selectionEnd;
      valeurNative(el, el.value.slice(0, d) + texte + el.value.slice(f));
      try { el.selectionStart = el.selectionEnd = d + texte.length; } catch (e) {}
      evenementSaisie(el, "insertText", texte);
      return true;
    }
    var sel = root.getSelection();
    if (sel && sel.rangeCount) {
      var r = sel.getRangeAt(0);
      r.deleteContents();
      var n = document.createTextNode(texte);
      r.insertNode(n);
      r.setStartAfter(n);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
    } else el.textContent += texte;
    evenementSaisie(el, "insertText", texte);
    return true;
  }
  function texteDe(el) { return estChamp(el) ? el.value : el.textContent || ""; }

  // Curseur au tout début du corps (avant une éventuelle signature).
  function curseurDebut(el) {
    var cible = el;
    while (cible.firstElementChild && !/^(BR|IMG|TABLE)$/.test(cible.firstElementChild.tagName)) cible = cible.firstElementChild;
    var r = document.createRange();
    r.setStart(cible, 0);
    r.collapse(true);
    var s = root.getSelection();
    s.removeAllRanges();
    s.addRange(r);
  }

  async function attendreQue(test, delaiMs, pasMs) {
    var fin = Date.now() + delaiMs * facteur();
    for (;;) {
      var v = test();
      if (v) return v;
      if (Date.now() > fin) return null;
      await attendre(pasMs || 200);
    }
  }

  // ------------------------------------------------------------ étapes
  async function ouvrirRedaction() {
    redaction = null; prevu = {};
    var b0 = await attendreFenetre(45000);
    if (b0) return { ok: false, raison: b0 };
    if (redactionOuverte()) return { ok: false, raison: "Un message est déjà ouvert en rédaction. Envoyez-le ou fermez-le, puis relancez." };
    var bouton = premier(SEL.nouveau) || parTexte(SEL.nouveauTexte);
    if (!bouton) return { ok: false, raison: "Bouton « Nouveau message » introuvable. Ouvrez la boîte de réception d'Outlook web." };
    // Ce qui existe déjà (brouillons repliés, réponses, champs cachés) n'est pas à nous.
    var anciens = tous(SEL.a).concat(tous(SEL.objet), tous(SEL.corps));
    await cliquer(bouton);
    var trouvee = await attendreQue(function () {
      var a = premierNouveau(SEL.a, anciens), o = premierNouveau(SEL.objet, anciens), c = premierNouveau(SEL.corps, anciens);
      return a && o && c ? conteneurDe(a, o, c) : null;
    }, 15000, 250);
    var b1 = await attendreFenetre(45000);
    if (b1) return { ok: false, raison: b1 };
    if (!trouvee) return { ok: false, raison: "La fenêtre de rédaction ne s'est pas ouverte (champs À, Objet, Corps et bouton Envoyer d'un même message introuvables)." };
    // Une seule nouvelle rédaction, sinon on ne sait pas laquelle est la nôtre.
    var autres = tous(SEL.objet).filter(function (o) { return anciens.indexOf(o) < 0 && visible(o) && !trouvee.contains(o); });
    if (autres.length) return { ok: false, raison: "Plusieurs nouvelles rédactions se sont ouvertes : rien n'est envoyé." };
    redaction = trouvee;
    await attendre(hasard(500, 1100));
    if (!redactionIntacte()) { redaction = null; return { ok: false, raison: PERDUE }; }
    return { ok: true };
  }

  // Pastilles de destinataires. Dans Outlook web récent, elles sont DANS le champ
  // « À » (élément non éditable, role="group", id « REK… », libellé « offline » + nom)
  // et montrent le nom du contact, pas forcément l'adresse.
  var SEL_PASTILLE = '[class*="_EType_RECIPIENT_ENTITY"],[id^="REK"],[contenteditable="false"],[role="group"],[role="listitem"],.ms-PickerPersona-container,[data-selection-index],[data-hovercard-id]';
  // La « ligne » d'un champ de destinataires (À ou Cci) : le champ et ses parents
  // proches, sans l'objet, le corps ni les autres champs de destinataires.
  function ligneDe(c) {
    if (!c) return null;
    var autres = [champObjet(), champCorps(), champA(), champCc(), champCci()].filter(function (x) { return x && x !== c; });
    var p = c;
    for (var i = 0; i < 3; i++) {
      var q = p.parentElement;
      if (!q || autres.some(function (x) { return q.contains(x); })) break;
      p = q;
    }
    return p;
  }
  function pastilles() { return pastillesDe(champA()); }
  function pastillesDe(c) {
    var z = ligneDe(c);
    if (!z) return [];
    return [].slice.call(z.querySelectorAll(SEL_PASTILLE)).filter(function (el) {
      if (el === c || el.contains(c) || !visible(el)) return false;
      if (!(el.textContent || "").trim() && !el.getAttribute("aria-label")) return false;
      if (el.closest('[role="listbox"],[role="option"]')) return false; // pas une suggestion
      // La plus extérieure seulement : on ignore les conteneurs du champ lui-même
      // (la ligne « À » peut porter role="group" sans être une pastille).
      for (var p = el.parentElement; p && p !== z; p = p.parentElement) {
        if (p === c || p.contains(c)) break;
        if (p.matches(SEL_PASTILLE)) return false;
      }
      return true;
    });
  }
  // Une pastille dont le libellé, l'info-bulle ou le texte montre une adresse
  // DIFFÉRENTE de celle(s) prévue(s) : Outlook a remplacé l'adresse par une autre
  // personne (autocomplétion). Une pastille qui ne montre qu'un nom ne dit rien.
  function autreAdresse(liste, prevues) {
    var attendues = [].concat(prevues).map(function (x) { return String(x).toLowerCase(); });
    // Le vrai Outlook colle l'état de présence au libellé (« offline » + nom ou adresse) :
    // on le retire, puis on ne lit que des adresses ENTIÈRES (bornées).
    var PRESENCE = /^\s*(offline|online|away|busy|available|do ?not ?disturb|dnd|out of office|hors connexion|en ligne|absent|occup[ée]|disponible|ne pas d[ée]ranger|inconnu|unknown)\s*/i;
    var ENTIERE = /(?<=^|[\s;,<>()\[\]"':=|])[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}(?=$|[\s;,<>()\[\]"':|.?!])/g;
    for (var i = 0; i < liste.length; i++) {
      var p = liste[i], morceaux = [p.getAttribute("aria-label"), p.getAttribute("title"), p.getAttribute("data-hovercard-id")];
      // Le texte, nœud par nœud (jamais recollé : « adresse » + « x » du bouton ≠ une adresse).
      var w = document.createTreeWalker(p, 4);
      for (var n = w.nextNode(); n; n = w.nextNode()) morceaux.push(n.nodeValue);
      [].forEach.call(p.querySelectorAll("[title],[aria-label],[data-hovercard-id]"), function (x) { morceaux.push(x.getAttribute("title"), x.getAttribute("aria-label"), x.getAttribute("data-hovercard-id")); });
      var vus = [];
      morceaux.forEach(function (m) { if (m) vus = vus.concat(String(m).replace(PRESENCE, "").replace(/[\u200B-\u200D\u2060\uFEFF]/g, "").match(ENTIERE) || []); });
      for (var k = 0; k < vus.length; k++) if (attendues.indexOf(vus[k].toLowerCase()) < 0) return vus[k];
    }
    return null;
  }
  // Texte saisi dans le champ, hors pastilles (ce qui n'a pas encore été validé).
  function texteLibre(c) {
    var out = "", w = document.createTreeWalker(c, 4); // NodeFilter.SHOW_TEXT
    for (var n = w.nextNode(); n; n = w.nextNode()) {
      var pastille = n.parentElement && n.parentElement.closest(SEL_PASTILLE);
      if (pastille && pastille !== c && c.contains(pastille)) continue;
      out += n.nodeValue;
    }
    return out;
  }
  // L'adresse apparaît-elle dans un attribut (libellé, info-bulle) d'une pastille ou de la rédaction ?
  function adresseValidee(adresse, portee, champ) {
    var zone = portee || zoneRedaction() || document;
    var cible = adresse.toLowerCase();
    var tous = zone.querySelectorAll("[aria-label],[title],[data-hovercard-id],span,div");
    for (var i = 0; i < tous.length; i++) {
      var el = tous[i];
      if (el === (champ || champA())) continue;
      if (el.closest && el.closest('[role="listbox"],[role="option"]')) continue;
      if (el === champObjet() || el === champCorps() || (champCorps() && champCorps().contains(el))) continue;
      var attrs = ((el.getAttribute("aria-label") || "") + " " + (el.getAttribute("title") || "") + " " + (el.getAttribute("data-hovercard-id") || "")).toLowerCase();
      if (attrs.indexOf(cible) >= 0) return true;
      if (!el.children.length && (el.textContent || "").trim().toLowerCase() === cible) return true;
    }
    return false;
  }

  // Coller comme Ctrl+V : un événement « paste » porteur du texte (et du HTML pour
  // le message), que le champ d'Outlook traite lui-même. S'il ne le prend pas,
  // insertion directe du même contenu.
  async function coller(el, html, texte) {
    el.focus();
    var gere;
    try {
      var dt = new DataTransfer();
      if (html) dt.setData("text/html", html);
      dt.setData("text/plain", texte);
      var ev = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true, composed: true });
      gere = !el.dispatchEvent(ev); // annulé = Outlook a pris le collage en charge
    } catch (e) { gere = false; }
    if (!gere) {
      if (!focusSur(el)) return false; // le curseur est ailleurs (fenêtre) : on ne colle pas chez elle
      var ok = false;
      if (html && !estChamp(el)) { try { ok = document.execCommand("insertHTML", false, html); } catch (e) { ok = false; } }
      if (!ok && !insererTexte(el, texte)) return false;
    }
    await attendre(hasard(400, 900)); // Outlook traite parfois le collage en différé
    return true;
  }
  function sansBarre(h) { return String(h || "").replace(/\/+$/, "").toLowerCase(); }
  function normaliser(t) { return String(t || "").replace(/[\s\u00a0]+/g, " ").trim(); }
  // Sans aucun espace : l'éditeur colle les paragraphes (« BonjourTexte ») là où le
  // texte brut met des retours à la ligne ; on compare donc sans les espaces.
  function compact(t) { return String(t || "").replace(/[\s\u00a0\u200B-\u200D\u2060\uFEFF]+/g, ""); }

  // « À » : l'adresse collée suivie de « ; », qu'Outlook transforme en pastille.
  // Validée quand le texte collé a quitté le champ et qu'une pastille est apparue
  // (portant l'adresse ou le nom du contact), ou que le champ est redevenu vide.
  // Texte restant hors pastilles, réduit aux caractères possibles d'une adresse :
  // Outlook glisse des caractères invisibles (espaces de largeur nulle) autour
  // des pastilles, qui ne doivent pas compter comme du texte non validé.
  function resteA() { return resteDe(champA()); }
  function resteDe(c) { return c ? texteLibre(c).replace(/[^A-Za-z0-9@._%+'-]/g, "") : ""; }
  function aValidee(adresse, avant) {
    if (!champA()) return false;
    var reste = resteA();
    if (reste) return false; // du texte d'adresse attend encore d'être validé
    return pastilles().length > avant || adresseValidee(adresse);
  }
  function etatA() { return etatChamp(champA()); }
  function etatChamp(c) {
    var libre = c ? texteLibre(c) : "";
    var visibles = libre.replace(/[\s\u200B-\u200D\u2060\uFEFF]/g, "");
    // Structure du champ, pour comprendre un échec sans avoir à deviner.
    var enfants = c ? [].slice.call(c.children).slice(0, 4).map(function (e) {
      return e.tagName.toLowerCase() + (e.id ? "#" + e.id : "") + (e.getAttribute("role") ? "[role=" + e.getAttribute("role") + "]" : "") +
        (e.getAttribute("contenteditable") ? "[ce=" + e.getAttribute("contenteditable") + "]" : "") + (typeof e.className === "string" && e.className ? "." + e.className.split(/\s+/).slice(0, 2).join(".") : "");
    }).join(" ") : "";
    var parents = [], p = c && c.parentElement;
    for (var i = 0; p && i < 3; i++, p = p.parentElement) parents.push(p.tagName.toLowerCase() + (p.getAttribute("role") ? "[role=" + p.getAttribute("role") + "]" : "") + (p.getAttribute("contenteditable") ? "[ce=" + p.getAttribute("contenteditable") + "]" : ""));
    return "pastilles vues : " + pastillesDe(c).length + ", texte restant : « " + visibles.slice(0, 60) + " » (" + libre.length + " caractère(s) en tout), enfants du champ : " + (enfants || "aucun") + ", parents : " + parents.join(" < ");
  }
  async function remplirA(adresse) {
    var b = await attendreFenetre(45000);
    if (b) return { ok: false, raison: b };
    if (!redactionIntacte()) return { ok: false, raison: PERDUE };
    var occ = await attendreLibre();
    if (occ) return { ok: false, raison: occ };
    var champ = champA();
    if (!champ) return { ok: false, raison: "Champ « À » introuvable." };
    var avant = pastilles().length;
    if (avant || resteA()) return { ok: false, raison: "Le champ « À » n'est pas vide : un seul destinataire par mail." };
    await cliquer(champ);
    await attendre(hasard(200, 500));
    if ((await coller(champA() || champ, null, adresse + ";")) === false) return { ok: false, raison: "Le curseur n'a pas pu être placé dans « À » (une fenêtre l'a pris) : rien n'est envoyé." };
    var ok = await attendreQue(function () { return aValidee(adresse, avant); }, 6000, 200);
    if (!ok) {
      // Collée mais pas validée : un « ; » de plus, comme on le ferait à la main.
      var c = champA();
      if (c && resteA().toLowerCase().indexOf(adresse.toLowerCase()) >= 0) {
        if (insererTexte(c, ";") === false) return { ok: false, raison: "Le curseur a quitté « À » : rien n'est envoyé." };
        ok = await attendreQue(function () { return aValidee(adresse, avant); }, 4000, 200);
      }
    }
    if (!ok) return { ok: false, raison: "L'adresse " + adresse + " n'a pas été validée dans « À » (" + etatA() + ")." };
    if (pastilles().length - avant > 1) return { ok: false, raison: "Plusieurs destinataires dans « À » : rien n'est envoyé." };
    var autre = autreAdresse(pastilles(), adresse);
    if (autre) return { ok: false, raison: "Outlook a mis « " + autre + " » au lieu de « " + adresse + " » dans « À » : rien n'est envoyé." };
    prevu.mode = "a"; prevu.n = 1; prevu.adresses = [adresse];
    await attendre(hasard(300, 800));
    return { ok: true };
  }

  // « À » vide (envoi en Cci) : ni pastille, ni texte en attente.
  function aVide() { var c = champA(); return !!c && !pastillesDe(c).length && !resteDe(c); }

  // Bouton qui fait apparaître le champ « Cci » (à côté de « À »).
  function boutonCci() {
    var zone = zoneRedaction() || document;
    var tous = zone.querySelectorAll('button,[role="button"],[role="link"],a');
    var proche = null;
    for (var i = 0; i < tous.length; i++) {
      var el = tous[i];
      if (!visible(el) || el.isContentEditable || (el.closest && el.closest('[class*="_EType_RECIPIENT_ENTITY"],[id^="REK"]'))) continue;
      var libs = [el.getAttribute("aria-label"), el.getAttribute("title"), el.textContent].map(function (x) { return (x || "").trim(); });
      if (libs.some(function (x) { return /^(cci|bcc)$/i.test(x); })) return el;
      if (!proche && libs.some(function (x) { return x.length < 60 && LIB.montrerCci.test(x); })) proche = el;
    }
    return proche;
  }

  // « Cci » : un paquet d'adresses collées d'un coup, séparées par « ; », que
  // Outlook transforme en pastilles. Validé quand il ne reste plus de texte à
  // valider et qu'il y a une pastille par adresse (les pastilles montrent souvent
  // le nom du contact, pas l'adresse). « À » doit rester vide.
  async function remplirCci(adresses) {
    var b = await attendreFenetre(45000);
    if (b) return { ok: false, raison: b };
    if (!redactionIntacte()) return { ok: false, raison: PERDUE };
    var occ = await attendreLibre();
    if (occ) return { ok: false, raison: occ };
    var n = adresses.length;
    if (!n) return { ok: false, raison: "Aucune adresse pour « Cci »." };
    var champ = champCci();
    if (!champ) {
      var bouton = boutonCci();
      if (!bouton) return { ok: false, raison: "Bouton « Cci » introuvable à côté de « À » (⚙ > Lancer le diagnostic)." };
      await cliquer(bouton);
      champ = await attendreQue(champCci, 5000, 200);
      if (!champ) return { ok: false, raison: "Le champ « Cci » ne s'est pas affiché après un clic sur « Cci »." };
      await attendre(hasard(300, 700));
    }
    var avant = pastillesDe(champ).length;
    if (avant || resteDe(champ)) return { ok: false, raison: "Le champ « Cci » n'est pas vide : rien n'est envoyé." };
    await cliquer(champ);
    await attendre(hasard(200, 500));
    if ((await coller(champCci() || champ, null, adresses.join("; ") + ";")) === false) return { ok: false, raison: "Le curseur n'a pas pu être placé dans « Cci » (une fenêtre l'a pris) : rien n'est envoyé." };
    var valide = function () {
      var c = champCci();
      if (!c || resteDe(c)) return false;
      if (pastillesDe(c).length - avant >= n) return true;
      var ligne = ligneDe(c);
      return adresses.every(function (a) { return adresseValidee(a, ligne, c); });
    };
    var ok = await attendreQue(valide, 6000 + 250 * n, 250);
    if (!ok) {
      // Collées mais pas toutes validées : un « ; » de plus, comme à la main.
      var c = champCci();
      if (c && resteDe(c)) {
        if (insererTexte(c, ";") === false) return { ok: false, raison: "Le curseur a quitté « Cci » : rien n'est envoyé." };
        ok = await attendreQue(valide, 4000, 200);
      }
    }
    if (!ok) return { ok: false, raison: n + " adresse(s) collée(s) dans « Cci » mais pas toutes validées (" + etatChamp(champCci()) + ")." };
    var vues = pastillesDe(champCci()).length - avant;
    if (vues > n) return { ok: false, raison: "Plus de destinataires que prévu dans « Cci » (" + vues + " au lieu de " + n + ") : rien n'est envoyé." };
    var autreCci = autreAdresse(pastillesDe(champCci()), adresses);
    if (autreCci) return { ok: false, raison: "Outlook a mis « " + autreCci + " » dans « Cci », qui n'était pas prévue : rien n'est envoyé." };
    if (!aVide()) return { ok: false, raison: "Le champ « À » n'est pas vide alors que l'envoi est en Cci : rien n'est envoyé." };
    prevu.mode = "cci"; prevu.n = n; prevu.adresses = adresses.slice();
    await attendre(hasard(300, 800));
    return { ok: true };
  }

  async function remplirObjet(objet) {
    var b = await attendreFenetre(45000);
    if (b) return { ok: false, raison: b };
    if (!redactionIntacte()) return { ok: false, raison: PERDUE };
    var occ = await attendreLibre();
    if (occ) return { ok: false, raison: occ };
    var champ = champObjet();
    if (!champ) return { ok: false, raison: "Champ « Objet » introuvable." };
    await cliquer(champ);
    await attendre(hasard(200, 500));
    if ((await coller(champObjet() || champ, null, objet)) === false) return { ok: false, raison: "Le curseur n'a pas pu être placé dans « Objet » (une fenêtre l'a pris) : rien n'est envoyé." };
    if (texteDe(champObjet() || champ).trim() !== String(objet).trim()) return { ok: false, raison: "L'objet collé ne correspond pas à l'objet prévu." };
    prevu.objet = String(objet).trim();
    await attendre(hasard(300, 800));
    return { ok: true };
  }

  // ------------------------------------------------------------ signature
  // Retirer le texte de la signature (nom, classe, ville…) en gardant ses images
  // (logo, réseaux). La signature, c'est : le texte déjà dans le corps avant le
  // collage du message, et le texte des blocs « Signature » d'Outlook (posés
  // parfois après coup par un complément comme Letsignit). Seuls des nœuds de
  // texte sont retirés : images, liens et mise en page restent.
  var signature = null; // { origine: [nœuds de texte], textes: [textes normalisés], brut, premiere }
  function noeudsTexte(el) {
    var out = [], w = document.createTreeWalker(el, 4); // NodeFilter.SHOW_TEXT
    for (var n = w.nextNode(); n; n = w.nextNode()) if (n.nodeValue.replace(/[\s\u200B-\u200D\u2060\uFEFF]/g, "")) out.push(n);
    return out;
  }
  function occurrences(texte, motif) {
    if (!motif) return 0;
    var n = 0;
    for (var i = texte.indexOf(motif); i >= 0; i = texte.indexOf(motif, i + motif.length)) n++;
    return n;
  }
  function retirerTexteSignature() {
    var corps = champCorps();
    if (!corps || !signature) return { ok: false, raison: "Corps du message introuvable." };
    var cibles = signature.origine.filter(function (n) { return n.isConnected && corps.contains(n); });
    [].forEach.call(corps.querySelectorAll('[id*="signature" i]'), function (bloc) {
      // Jamais le bloc où le message aurait été collé.
      if (signature.premiere && normaliser(bloc.textContent).indexOf(signature.premiere) >= 0) return;
      noeudsTexte(bloc).forEach(function (n) { if (cibles.indexOf(n) < 0) cibles.push(n); });
    });
    if (!cibles.length) {
      // L'éditeur a reconstruit le corps : on retrouve les textes de la signature
      // en partant de la fin (la signature est en bas), sans toucher à ceux du message.
      var texteCorps = normaliser(corps.textContent), texteMessage = normaliser(signature.brut);
      var surplus = {};
      signature.textes.forEach(function (t) { surplus[t] = occurrences(texteCorps, t) - occurrences(texteMessage, t); });
      noeudsTexte(corps).reverse().forEach(function (n) {
        var t = normaliser(n.nodeValue);
        if (surplus[t] > 0) { cibles.push(n); surplus[t]--; }
      });
    }
    cibles.forEach(function (n) { n.remove(); });
    if (cibles.length) evenementSaisie(corps, "deleteContentBackward", null);
    // Contrôles : le message est intact, et aucun texte de la signature ne reste.
    var apres = normaliser(corps.textContent);
    if (signature.premiere && apres.indexOf(signature.premiere) < 0) return { ok: false, raison: "Le message a été touché en retirant la signature : rien n'est envoyé." };
    var restant = signature.textes.filter(function (t) { return occurrences(apres, t) > occurrences(normaliser(signature.brut), t); })[0];
    if (restant) return { ok: false, raison: "Le texte de la signature (« " + restant.slice(0, 40) + " ») n'a pas pu être retiré : rien n'est envoyé." };
    return { ok: true, retires: cibles.length };
  }

  // Message : collé d'un coup au début du corps (avant la signature), mise en forme
  // et liens compris. Vérifie que le texte et chaque lien sont bien dans le mail.
  // options.sansTexteSignature : retire ensuite le texte de la signature (images gardées).
  async function remplirCorps(html, options) {
    var M = PP.Message;
    var b = await attendreFenetre(45000);
    if (b) return { ok: false, raison: b };
    if (!redactionIntacte()) return { ok: false, raison: PERDUE };
    var occ = await attendreLibre();
    if (occ) return { ok: false, raison: occ };
    var champ = champCorps();
    if (!champ) return { ok: false, raison: "Corps du message introuvable." };
    var propre = M.nettoyer(html);
    var brut = M.texteBrut(propre);
    await cliquer(champ);
    champ.focus();
    curseurDebut(champ);
    var avantTexte = normaliser(champ.textContent); // la signature, en général
    var origine = noeudsTexte(champ);
    // Le nouveau message ne doit pas contenir le texte du message précédent (reste d'un brouillon réutilisé).
    if (dernierMessage && compact(champ.textContent).indexOf(dernierMessage) >= 0) return { ok: false, raison: "Le nouveau message contient déjà le texte d'un message précédent : rien n'est envoyé." };
    if ((await coller(champ, propre, brut)) === false) return { ok: false, raison: "Le curseur n'a pas pu être placé dans le corps (une fenêtre l'a pris) : rien n'est envoyé." };
    await attendre(hasard(500, 1200));
    var corps = champCorps() || champ;
    // Garde-fou : rien d'autre que le message ne doit être entré (par exemple le
    // contenu de votre presse-papiers, si Outlook le lisait de lui-même).
    var ajoute = normaliser(corps.textContent).length - avantTexte.length;
    var attendu = normaliser(brut).length;
    if (ajoute > attendu + Math.max(40, attendu * 0.2)) return { ok: false, raison: "Le mail contient plus de texte que le message prévu (presse-papiers collé par Outlook ?) : rien n'est envoyé. Dans Chrome, bloquez l'accès d'Outlook au presse-papiers." };
    var premiere = normaliser(brut.split(/\r?\n/).filter(function (l) { return l.trim(); })[0] || "").slice(0, 30);
    if (premiere && normaliser(corps.textContent).indexOf(premiere) < 0) return { ok: false, raison: "Le texte n'apparaît pas dans le corps du message." };
    var presents = [].slice.call(corps.querySelectorAll("a[href]")).map(function (a) { return sansBarre(a.getAttribute("href")); });
    var manquants = M.liens(propre).filter(function (l) { return presents.indexOf(sansBarre(l.href)) < 0; });
    if (manquants.length) return { ok: false, raison: "Lien absent du mail : " + manquants[0].href };
    signature = { origine: origine, textes: origine.map(function (n) { return normaliser(n.nodeValue); }), brut: brut, premiere: premiere };
    prevu.message = compact(brut); prevu.liens = M.liens(propre).map(function (l) { return sansBarre(l.href); });
    prevu.longueurMax = avantTexte.length + attendu + Math.max(40, attendu * 0.2);
    if (options && options.sansTexteSignature) {
      var r = retirerTexteSignature();
      if (!r.ok) return r;
    }
    return { ok: true };
  }

  // ------------------------------------------------------------ importance
  // « Importance haute » : bouton à bascule du ruban (onglets « Message » ou
  // « Options »), bouton « Importance » à menu, ou menu « Plus d'options » >
  // « Définir l'importance » > « Haute ». Ne clique que sur des commandes dont le
  // libellé parle d'importance, jamais deux fois sur la même, et ne compte comme
  // réussi que ce qu'Outlook montre ensuite activé (aria-pressed, aria-checked) :
  // sinon rien n'est envoyé. Chaque étape est notée (journal du diagnostic).
  var SEL_COMMANDE = 'button,[role="button"],[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"],[role="option"],[role="tab"],[role="checkbox"],[role="switch"],[role="combobox"]';
  var SEL_MENU = '[role="menu"],.ms-ContextualMenu,[role="menuitem"],[role="menuitemcheckbox"],[role="menuitemradio"]';
  function libelles(el) {
    return [el.getAttribute("aria-label"), el.getAttribute("title"), el.textContent].map(function (x) { return (x || "").trim().replace(/\s+/g, " "); }).filter(Boolean);
  }
  // Commandes visibles dont un libellé correspond. dansMenu : true = dans un menu
  // ouvert (ou option de liste déroulante), false = ailleurs (ruban, rédaction).
  // Jamais dans la liste des messages ni des dossiers.
  function commandes(motif, dansMenu) {
    return [].slice.call(document.querySelectorAll(SEL_COMMANDE)).filter(function (el) {
      if (!visible(el) || el.closest('[role="grid"],[role="tree"],[role="treegrid"]')) return false;
      var menu = !!el.closest(SEL_MENU) || el.getAttribute("role") === "option";
      if (dansMenu != null && menu !== dansMenu) return false;
      return libelles(el).some(function (x) { return x.length < 60 && motif.test(x); });
    });
  }
  function active(el) {
    var v = el.getAttribute("aria-pressed") || el.getAttribute("aria-checked");
    return v === "true" ? true : v === "false" ? false : null;
  }
  function menuOuvert() { return [].some.call(document.querySelectorAll('[role="menu"],.ms-ContextualMenu'), visible); }
  async function fermerMenus() {
    for (var i = 0; i < 3 && menuOuvert(); i++) {
      var cible = document.activeElement || document.body;
      ["keydown", "keyup"].forEach(function (type) {
        cible.dispatchEvent(new KeyboardEvent(type, { key: "Escape", code: "Escape", keyCode: 27, which: 27, bubbles: true, cancelable: true, composed: true }));
      });
      await attendre(300);
    }
  }
  var journal = [];
  function noter(t) { journal.push(t); }
  function nomCommande(el) {
    var p = el.getAttribute("aria-pressed"), c = el.getAttribute("aria-checked");
    return (el.getAttribute("role") || el.tagName.toLowerCase()) + " « " + (libelles(el)[0] || "").slice(0, 50) + " »" +
      (p != null ? " [aria-pressed=" + p + "]" : c != null ? " [aria-checked=" + c + "]" : " [sans état]");
  }
  // Un bouton à bascule « Importance haute » est-il affiché enfoncé ?
  function hauteActivee() { return commandes(LIB.importanceHaute, false).some(function (el) { return active(el) === true; }); }
  // Dans le menu ouvert : « Haute » (directement, ou dans le sous-menu « Définir l'importance »).
  async function itemHaute() {
    var trouver = function () {
      return commandes(LIB.choixHaute, true)[0] || commandes(LIB.importanceHaute, true).filter(function (el) { return el.getAttribute("role") !== "option"; })[0];
    };
    var item = trouver();
    if (item) return item;
    var sous = commandes(LIB.importance, true)[0];
    if (!sous) return null;
    noter("sous-menu : " + nomCommande(sous));
    await cliquer(sous);
    await attendre(hasard(400, 800));
    return trouver();
  }
  var cliques = []; // commandes à bascule déjà cliquées dans cette rédaction
  // Bouton à bascule visible. null : aucun ; true : activé (vu enfoncé) ; false : pas confirmé.
  async function parBouton() {
    var b = commandes(LIB.importanceHaute, false).filter(function (el) { return cliques.indexOf(el) < 0; })[0];
    if (!b) return null;
    noter("bouton : " + nomCommande(b));
    if (active(b) === true) return true;
    cliques.push(b);
    await cliquer(b);
    await attendre(hasard(500, 900));
    if (menuOuvert()) { // bouton à menu déroulant
      var item = await itemHaute();
      noter("son menu : " + (item ? nomCommande(item) : "pas de « Haute »"));
      if (item && active(item) !== true) { await cliquer(item); await attendre(hasard(400, 800)); }
      await fermerMenus();
    }
    var apres = commandes(LIB.importanceHaute, false)[0];
    noter("après le clic : " + (apres ? nomCommande(apres) : "bouton disparu"));
    return hauteActivee();
  }
  // Menu (bouton « Importance » ou « Plus d'options ») : choisir « Haute », puis
  // rouvrir le menu pour lire son état. null : pas de « Haute » dans ce menu.
  async function parMenu(trouverOuvreur) {
    var o = trouverOuvreur();
    if (!o) return null;
    noter("menu : " + nomCommande(o));
    await cliquer(o);
    await attendre(hasard(500, 900));
    var item = await itemHaute();
    if (!item) { noter("pas de « Haute » dans ce menu"); await fermerMenus(); return null; }
    noter("choix : " + nomCommande(item));
    if (active(item) !== true) { await cliquer(item); await attendre(hasard(400, 800)); }
    await fermerMenus();
    if (hauteActivee()) return true;
    o = trouverOuvreur();
    if (!o) return false;
    await cliquer(o);
    await attendre(hasard(500, 900));
    var verif = await itemHaute();
    noter("menu rouvert : " + (verif ? nomCommande(verif) : "« Haute » introuvable"));
    var ok = !!verif && active(verif) === true;
    await fermerMenus();
    return ok;
  }
  async function importanceHaute() {
    journal = [];
    cliques = [];
    var b = await attendreFenetre(45000);
    if (b) return { ok: false, raison: b };
    var r = null, vu = false;
    var tenter = async function (essai) {
      if (r === true) return;
      var x = await essai();
      if (x !== null) vu = true;
      if (x === true) r = true;
    };
    var menuImportance = function () { return parMenu(function () { return commandes(LIB.importance, false)[0]; }); };
    // Dans l'onglet affiché : bouton à bascule, menu « Importance », puis menus
    // « Plus d'options » / « Autres options » (« … ») de la rédaction et du ruban.
    var essaisOnglet = async function () {
      await tenter(parBouton);
      await tenter(menuImportance);
      var zone = zoneRedaction();
      var plus = r === true ? [] : commandes(LIB.plusOptions, false).sort(function (x, y) { return (zone && zone.contains(y) ? 1 : 0) - (zone && zone.contains(x) ? 1 : 0); });
      for (var i = 0; r !== true && i < plus.length && i < 4; i++) {
        var p = plus[i];
        await tenter(function () { return parMenu(function () { return p.isConnected && visible(p) ? p : null; }); });
      }
    };
    await essaisOnglet();
    // Onglets du ruban de la rédaction où se trouve souvent la commande.
    var onglets = [LIB.ongletMessage, LIB.ongletOptions];
    for (var k = 0; r !== true && k < onglets.length; k++) {
      var onglet = commandes(onglets[k], false).filter(function (el) { return el.getAttribute("role") === "tab"; })[0];
      if (!onglet || onglet.getAttribute("aria-selected") === "true") continue;
      noter("onglet : " + nomCommande(onglet));
      await cliquer(onglet);
      await attendre(hasard(500, 900));
      await essaisOnglet();
    }
    if (r === true) { noter("importance haute confirmée"); return { ok: true }; }
    var resume = journal.slice(-3).join(" ; ");
    if (vu) return { ok: false, raison: "Outlook n'affiche pas « Importance haute » activée après le clic : rien n'est envoyé (" + resume + "). ⚙ > Lancer le diagnostic." };
    return { ok: false, raison: "Commande « Importance haute » introuvable dans Outlook : rien n'est envoyé" + (resume ? " (" + resume + ")" : "") + ". ⚙ > Lancer le diagnostic." };
  }

  // Clique « Envoyer » puis attend la fermeture de la rédaction. Toute autre issue
  // est « à vérifier » : le moteur s'arrête et ne renvoie jamais tout seul.
  // Dernier contrôle, juste avant « Envoyer » : la rédaction contient EXACTEMENT ce
  // que Smart Mail y a mis (rien n'a été modifié entre-temps : frappe de
  // l'utilisateur, complément, réaffichage d'Outlook). Sinon rien ne part.
  function controleAvantEnvoi() {
    if (!redactionIntacte()) return PERDUE;
    if (!prevu.mode || prevu.objet == null || prevu.message == null) return "Le message n'a pas été entièrement préparé : rien n'est envoyé.";
    if (prevu.souille) return "Une frappe au clavier a atteint le message en préparation : rien n'est envoyé.";
    var a = champA(), cc = champCc(), cci = champCci(), o = champObjet(), c = champCorps();
    if (texteDe(o).trim() !== prevu.objet) return "L'objet a changé depuis sa saisie : rien n'est envoyé.";
    if (cc && (pastillesDe(cc).length || resteDe(cc))) return "Le champ « Cc » n'est pas vide : rien n'est envoyé.";
    if (prevu.mode === "a") {
      if (pastillesDe(a).length !== 1 || resteDe(a)) return "Le champ « À » ne contient plus exactement le destinataire prévu : rien n'est envoyé.";
      if (cci && (pastillesDe(cci).length || resteDe(cci))) return "Le champ « Cci » n'est pas vide : rien n'est envoyé.";
    } else {
      if (!aVide()) return "Le champ « À » n'est pas vide alors que l'envoi est en Cci : rien n'est envoyé.";
      if (!cci || pastillesDe(cci).length !== prevu.n || resteDe(cci)) return "Le champ « Cci » ne contient plus exactement les " + prevu.n + " destinataires prévus : rien n'est envoyé.";
    }
    var texte = c.textContent || "";
    if (compact(texte).indexOf(prevu.message) < 0) return "Le message a été modifié depuis son collage : rien n'est envoyé.";
    if (normaliser(texte).length > prevu.longueurMax + 1500) return "Le mail contient beaucoup plus de texte que prévu : rien n'est envoyé.";
    var presents = [].slice.call(c.querySelectorAll("a[href]")).map(function (x) { return sansBarre(x.getAttribute("href")); });
    var manque = prevu.liens.filter(function (h) { return presents.indexOf(h) < 0; })[0];
    if (manque) return "Lien absent du mail : " + manque + " ; rien n'est envoyé.";
    return null;
  }

  var ECHEC_RE = /n.a pas pu .tre envoy|non envoy|not (been )?sent|could ?n.t be sent|could not be sent|failed to send|[ée]chec de l.envoi|nicht gesendet|konnte nicht gesendet|no se (pudo|ha podido) enviar|non è stato possibile inviare|kan niet worden verzonden|não foi possível enviar/i;
  function echecEnvoi() {
    var zones = document.querySelectorAll('[role="alert"],[role="alertdialog"],[role="status"],[role="dialog"],.ms-MessageBar,[class*="notification" i],[class*="toast" i]');
    for (var i = 0; i < zones.length; i++) {
      var t = (zones[i].textContent || "").replace(/\s+/g, " ").trim();
      if (visible(zones[i]) && ECHEC_RE.test(t)) return t.slice(0, 120);
    }
    return null;
  }
  // Bandeau d'envoi différé (« Annuler l'envoi », « Undo »).
  var ANNULER_RE = /annuler l.envoi|^annuler$|undo( send)?|envoi dans \d|sending in \d|r[ée]tablir/i;
  function annulationEnCours() {
    var zones = document.querySelectorAll('[role="alert"],[role="status"],.ms-MessageBar,[class*="notification" i],[class*="toast" i],[class*="undo" i]');
    for (var i = 0; i < zones.length; i++) {
      if (!visible(zones[i])) continue;
      var t = (zones[i].textContent || "").replace(/\s+/g, " ").trim();
      if (t.length < 160 && ANNULER_RE.test(t)) return true;
      var bs = zones[i].querySelectorAll('button,[role="button"]');
      for (var k = 0; k < bs.length; k++) if (ANNULER_RE.test((bs[k].getAttribute("aria-label") || bs[k].textContent || "").trim())) return true;
    }
    return false;
  }
  // Message explicite de blocage dans la rédaction (stratégie, étiquette obligatoire…).
  var REFUS_RE = /ne peut pas .tre envoy|ne pouvez pas envoyer|bloqu|interdit|can.?not be sent|can.?t send|blocked|not allowed|requires? a (sensitivity )?label|[ée]tiquette.*(obligatoire|requise)/i;
  function refusEnvoi() {
    var zones = (redaction || document).querySelectorAll('[role="alert"],[role="status"],.ms-MessageBar,[class*="policy" i],[class*="mailtip" i]');
    for (var i = 0; i < zones.length; i++) {
      var t = (zones[i].textContent || "").replace(/\s+/g, " ").trim();
      if (visible(zones[i]) && REFUS_RE.test(t)) return t.slice(0, 120);
    }
    return null;
  }
  function brouillonRevenu() {
    if (!prevu.objet) return null;
    var objets = tous(SEL.objet).filter(visible);
    for (var i = 0; i < objets.length; i++) if (texteDe(objets[i]).trim() === prevu.objet) return "brouillon « " + prevu.objet.slice(0, 40) + " » rouvert";
    return null;
  }
  async function envoyer() {
    var b = await attendreFenetre(45000);
    if (b) return { ok: false, envoye: false, raison: b };
    var occ = await attendreLibre();
    if (occ) return { ok: false, envoye: false, raison: occ };
    var probleme = controleAvantEnvoi();
    if (probleme) return { ok: false, envoye: false, raison: probleme };
    // Le bouton « Envoyer » de NOTRE rédaction, jamais un autre de la page.
    var bouton = premier(SEL.envoyer, redaction) || parTexte(SEL.envoyerTexte, redaction);
    if (!bouton) return { ok: false, envoye: false, raison: "Bouton « Envoyer » introuvable dans la rédaction : rien n'est envoyé." };
    if (bouton.disabled || bouton.getAttribute("aria-disabled") === "true") return { ok: false, envoye: false, raison: "Outlook a désactivé « Envoyer » pour ce message : rien n'est envoyé." };
    if (root.navigator && root.navigator.onLine === false) return { ok: false, envoye: false, raison: "Hors connexion : rien n'est envoyé." };
    var notre = redaction;
    var clique = await cliquer(bouton);
    if (clique === "absent") {
      // Disparue avant tout appui sur « Envoyer » : aucun geste, rien n'est parti.
      redaction = null;
      return { ok: false, envoye: false, raison: "La rédaction a disparu juste avant le clic sur « Envoyer » : rien n'est parti." };
    }
    if (clique === "incertain") {
      // Disparue PENDANT l'appui : on ne sait pas. Jamais renvoyé automatiquement.
      redaction = null;
      return { ok: false, envoye: null, raison: "La rédaction a disparu pendant le clic sur « Envoyer » : à vérifier, rien ne sera renvoyé automatiquement." };
    }
    // Un complément peut agir à l'envoi (vérification, signature) : jusqu'à 60 s.
    // Envoyé = NOTRE rédaction s'est refermée (une autre peut rester ouverte).
    var ferme = await attendreQue(function () {
      return !notre.isConnected || !visible(notre) || !(premier(SEL.a, notre) || premier(SEL.objet, notre));
    }, 60000, 250);
    if (ferme) {
      redaction = null;
      // Preuve d'envoi : quelques secondes sans message d'échec, et le brouillon
      // n'est pas revenu (Outlook rouvre un message qui n'a pas pu partir).
      var echec = await attendreQue(function () { return echecEnvoi() || brouillonRevenu(); }, 3000, 250);
      // « Annuler l'envoi » : le mail ne part qu'à la fin du compte à rebours.
      if (!echec && annulationEnCours()) {
        await attendreQue(function () { return !annulationEnCours() || echecEnvoi() || brouillonRevenu(); }, 60000, 250);
        echec = echecEnvoi() || brouillonRevenu() || (annulationEnCours() ? "délai d'annulation toujours affiché" : null);
      }
      if (echec) return { ok: false, envoye: null, raison: "Après « Envoyer », Outlook signale un problème ou a rouvert le brouillon (" + echec + ") : à vérifier, rien ne sera renvoyé automatiquement." };
      dernierMessage = prevu.message || "";
      return { ok: true };
    }
    // Rédaction restée ouverte : si Outlook dit clairement que rien n'est parti
    // (refus, blocage, désactivé), c'est certain ; sinon l'issue est incertaine.
    if (redactionIntacte()) {
      var refus = echecEnvoi() || refusEnvoi();
      if (refus) return { ok: false, envoye: false, raison: "Outlook a refusé l'envoi (« " + refus + " ») : rien n'est parti." };
    }
    var dialogue = dialogueBloquant() || premier(['[role="alertdialog"]', '[role="dialog"]']);
    var texteDialogue = dialogue ? (dialogue.textContent || "").trim().replace(/\s+/g, " ").slice(0, 200) : "";
    return { ok: false, envoye: null, raison: "Après « Envoyer », la rédaction est restée ouverte" + (texteDialogue ? " : " + texteDialogue : ".") };
  }

  // Diagnostic : ouvre un nouveau message, relève les champs (sans rien taper), referme.
  function decrire(el) {
    if (!el) return null;
    return { tag: el.tagName, role: el.getAttribute("role"), ariaLabel: el.getAttribute("aria-label"), placeholder: el.getAttribute("placeholder"),
      title: el.getAttribute("title"), editable: !!el.isContentEditable, html: (el.outerHTML || "").slice(0, 600) };
  }
  function balayage() {
    var out = [];
    var tous = document.querySelectorAll('input,[contenteditable="true"],[role="textbox"],[role="dialog"],button[aria-label],[role="button"][aria-label]');
    for (var i = 0; i < tous.length && out.length < 200; i++) if (visible(tous[i])) out.push(decrire(tous[i]));
    return out;
  }
  async function diagnostic() {
    redaction = null; prevu = {};
    var dialogues = function () {
      return [].slice.call(document.querySelectorAll('[role="dialog"],[role="alertdialog"],.ms-Modal.is-open')).filter(visible)
        .map(function (d) { var x = decrireDialogue(d); x.modale = d.getAttribute("aria-modal"); x.bloquante = d === dialogueBloquant(); x.html = (d.outerHTML || "").slice(0, 1500); return x; });
    };
    var rapport = { version: 4, quand: new Date().toString(), url: location.href.slice(0, 120),
      avant: { nouveau: decrire(premier(SEL.nouveau) || parTexte(SEL.nouveauTexte)), redactionOuverte: redactionOuverte(), dialogues: dialogues() } };
    // Une rédaction déjà ouverte est celle de l'utilisateur : on la décrit sans
    // JAMAIS cliquer dedans. Sinon on ouvre la nôtre (ancrée) et on la referme.
    var notre = false;
    if (!redactionOuverte() && !dialogueBloquant()) {
      var b = premier(SEL.nouveau) || parTexte(SEL.nouveauTexte);
      if (b) {
        var anciens = tous(SEL.a).concat(tous(SEL.objet), tous(SEL.corps));
        await cliquer(b);
        var c = await attendreQue(function () {
          var a = premierNouveau(SEL.a, anciens), o = premierNouveau(SEL.objet, anciens), co = premierNouveau(SEL.corps, anciens);
          return a && o && co ? conteneurDe(a, o, co) : null;
        }, 10000, 250);
        if (c) { redaction = c; notre = true; }
        await attendre(800);
      }
    }
    rapport.redactionDeLUtilisateur = !notre && redactionOuverte();
    var zone = zoneRedaction();
    rapport.redaction = { a: decrire(champA()), objet: decrire(champObjet()), corps: decrire(champCorps()),
      envoyer: decrire(zone && (premier(SEL.envoyer, zone) || parTexte(SEL.envoyerTexte, zone))), ignorer: decrire(premier(SEL.ignorer)), actif: decrire(document.activeElement),
      boutonCci: decrire(boutonCci()), signature: champCorps() ? (champCorps().innerHTML || "").slice(0, 3000) : null };
    // « Cci » : le faire apparaître (sans rien taper) pour relever le champ.
    if (notre && redactionIntacte() && !champCci() && boutonCci() && !dialogueBloquant()) { await cliquer(boutonCci()); await attendreQue(champCci, 3000, 200); }
    rapport.redaction.cci = decrire(champCci());
    rapport.redaction.ligneCci = champCci() ? (ligneDe(champCci()).outerHTML || "").slice(0, 1500) : null;
    // Importance : commandes qui en parlent, onglets du ruban, menus « Plus d'options ».
    var relever = function (motif) { return commandes(motif).slice(0, 12).map(function (el) { var d = decrire(el); d.etat = active(el); d.dansMenu = !!el.closest(SEL_MENU); return d; }); };
    var ongletsVus = function () { return commandes(/./).filter(function (el) { return el.getAttribute("role") === "tab"; }).slice(0, 20).map(function (el) { return libelles(el)[0] + (el.getAttribute("aria-selected") === "true" ? " (actif)" : ""); }); };
    rapport.importance = { commandes: relever(/importan/i), onglets: ongletsVus(), plus: relever(LIB.plusOptions) };
    var plus = notre ? commandes(LIB.plusOptions, false).filter(function (el) { return zone && zone.contains(el); })[0] : null;
    if (plus && !dialogueBloquant()) {
      await cliquer(plus);
      await attendre(800);
      rapport.importance.menuPlus = commandes(/./, true).slice(0, 40).map(function (el) { return libelles(el)[0]; });
      await fermerMenus();
    }
    // Ce que voit chaque onglet « Message » / « Options » du ruban (commandes courtes).
    rapport.importance.parOnglet = {};
    for (var k = 0; notre && k < 2 && !dialogueBloquant(); k++) {
      var onglet = commandes([LIB.ongletMessage, LIB.ongletOptions][k], false).filter(function (el) { return el.getAttribute("role") === "tab"; })[0];
      if (!onglet) continue;
      await cliquer(onglet);
      await attendre(900);
      rapport.importance.parOnglet[libelles(onglet)[0]] = {
        importance: relever(/importan/i),
        commandes: commandes(/^.{1,40}$/, false).filter(function (el) { return !zone || !zone.contains(el); }).slice(0, 80).map(function (el) { return libelles(el)[0] + (active(el) != null ? " [" + active(el) + "]" : ""); }),
      };
    }
    // Essai réel (dans ce brouillon, qui sera ignoré) : quelles commandes sont cliquées.
    if (notre && redactionIntacte() && !dialogueBloquant()) {
      var essai = await importanceHaute();
      rapport.importance.essai = { resultat: essai, journal: journal.slice() };
    }
    rapport.dialogues = dialogues();
    rapport.balayage = balayage();
    // Refermer NOTRE rédaction de diagnostic (jamais celle de l'utilisateur).
    var ign = notre && redactionIntacte() ? premier(SEL.ignorer, redaction) : null;
    if (ign && !dialogueBloquant()) {
      await cliquer(ign);
      await attendre(800);
      var dlg = premier(['[role="alertdialog"]', '[role="dialog"]']);
      var conf = dlg ? parTexte(["ok", "ignorer", "abandonner", "discard", "supprimer"], dlg) : null;
      if (conf && redaction && redaction.isConnected) await cliquer(conf);
    }
    rapport.refermee = notre ? !(redaction && redaction.isConnected && visible(redaction)) : null;
    redaction = null;
    return rapport;
  }

  // Pour le relevé d'adresses (mega) : où a-t-il le droit de taper ?
  function correspond(el, selecteurs) {
    for (var i = 0; i < selecteurs.length; i++) { try { if (el.matches(selecteurs[i])) return true; } catch (e) {} }
    return false;
  }
  // Un champ destinataire (À, Cc, Cci) d'une rédaction — jamais l'objet, le corps,
  // la recherche ou une autre zone de saisie.
  function estChampDestinataire(el) {
    return !!el && el.nodeType === 1 && (correspond(el, SEL.a) || correspond(el, SEL.cc) || correspond(el, SEL.cci));
  }
  // Ni pastille ni texte en attente.
  function champVide(el) { return !!el && !pastillesDe(el).length && !resteDe(el) && !texteDe(el).replace(/[\s\u200B-\u200D\u2060\uFEFF]/g, ""); }
  // La rédaction affichée est-elle vierge (aucun destinataire, aucun objet) ?
  // Sinon c'est un message de l'utilisateur : le relevé n'y touche pas.
  function redactionVide() {
    var a = premier(SEL.a), cc = premier(SEL.cc), cci = premier(SEL.cci), o = premier(SEL.objet);
    if (a && !champVide(a)) return false;
    if (cc && !champVide(cc)) return false;
    if (cci && !champVide(cci)) return false;
    if (o && texteDe(o).trim()) return false;
    return true;
  }

  PP.Outlook = {
    surAttente: null, // rappel posé par le moteur pour afficher les attentes
    MAIL: MAIL,
    attendre: attendre,
    hasard: hasard,
    redactionOuverte: redactionOuverte,
    champA: function () { return premier(SEL.a); }, // Smart Mail : « À » de la rédaction affichée (relevé d'adresses)
    estChampDestinataire: estChampDestinataire,
    champVide: champVide,
    redactionVide: redactionVide,
    ouvrirRedaction: ouvrirRedaction,
    remplirA: remplirA,
    remplirCci: remplirCci,
    aVide: aVide,
    remplirObjet: remplirObjet,
    remplirCorps: remplirCorps,
    retirerTexteSignature: retirerTexteSignature,
    importanceHaute: importanceHaute,
    envoyer: envoyer,
    diagnostic: diagnostic,
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
