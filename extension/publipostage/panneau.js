// Panneau flottant dans l'onglet Outlook web (Shadow DOM, styles isolés), même
// design que l'extension « Adresses ». Préparation (destinataires, liste « Déjà
// envoyés », message avec liens, rythme et Cci), vue compacte pendant l'envoi,
// suivi et export.

(function (root) {
  "use strict";
  var PP = root.PP || {};
  var api = root.chrome || root.browser;
  var Moteur = PP.Moteur, O = PP.Outlook, M = PP.Message, Modele = PP.Modele || { objet: "", html: "" };
  if (!Moteur || !O || !M) return;
  if (root.top !== root.self) return; // cadre principal uniquement

  var CLE_BROUILLON = "pp-brouillon";
  var CLE_IMPORT = "pp-import";
  var TAILLE_MAX = 200 * 1024 * 1024; // fichier importé : 200 Mo

  var CSS = `
  :host { all: initial; }
  * { box-sizing: border-box; }
  .panel { --bg:#F6F8FC;--card:#FFF;--line:#E6EBF2;--ink:#0F172A;--ink2:#334155;--muted:#5B6A80;--faint:#94A3B8;
    --brand:#1E4D8C;--brand-dark:#173E73;--brand-soft:#E8F0FA;--info-bg:#EEF4FC;--ok-bg:#DCFCE7;--ok:#15803D;
    --warn-bg:#FEF3C7;--warn:#92400E;--track:#E8EEF6;
    width:400px;max-width:92vw;max-height:88vh;overflow:auto;background:var(--card);color:var(--ink);
    border:1px solid var(--line);border-radius:14px;box-shadow:0 10px 40px rgba(15,23,42,.28);
    font:13px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; }
  @media (prefers-color-scheme: dark) {
    .panel { --bg:#0F141C;--card:#161D28;--line:#263042;--ink:#E7ECF3;--ink2:#C5CEDA;--muted:#93A1B5;--faint:#64748B;
      --brand:#4F86D1;--brand-dark:#3B72BE;--brand-soft:#1C2B42;--info-bg:#17253A;--ok-bg:#12301F;--ok:#5FD38D;
      --warn-bg:#33270F;--warn:#F2C46B;--track:#222C3B; } }
  .head { display:flex;align-items:center;gap:10px;padding:12px 14px;border-bottom:1px solid var(--line);cursor:grab;user-select:none; }
  .logo { width:30px;height:30px;border-radius:8px;background:linear-gradient(45deg,#102A5C,#3B82F6);color:#fff;display:grid;place-items:center;font-weight:800;font-size:11.5px;letter-spacing:-.2px;flex:none;position:relative;font-family:'Avenir Next','Segoe UI',system-ui,sans-serif; }
  .logo::after { content:'✦';position:absolute;top:-1px;right:2px;font-size:8px;color:#FBBF24;line-height:1; }
  .head h1 { font-size:15px;margin:0; } .head .sub { font-size:11px;color:var(--muted); }
  .version { font-size:11px;font-weight:600;color:var(--brand);background:var(--brand-soft);border-radius:6px;padding:1px 6px;margin-left:4px;vertical-align:2px; }
  .perime { background:var(--warn-bg);color:var(--warn);padding:9px 14px;font-size:12px;font-weight:600; }
  .iconbtn { border:0;background:transparent;color:var(--muted);font:inherit;font-size:16px;cursor:pointer;padding:2px 8px;border-radius:7px; }
  .iconbtn:hover { background:var(--bg); }
  .body { padding:14px; } .card { background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:10px; }
  .spread { display:flex;justify-content:space-between;align-items:center;gap:10px; }
  .row { display:flex;align-items:center;gap:8px; }
  .label { font-weight:600;color:var(--ink2);margin:0 0 6px; } .small { font-size:12px;color:var(--muted); }
  .help { font-size:12px;color:var(--muted);margin:6px 0 0; }
  .pills { display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 8px; }
  .pill { border:1px solid var(--line);background:var(--card);color:var(--ink2);font:inherit;font-weight:600;padding:6px 12px;border-radius:999px;cursor:pointer; }
  .pill[aria-pressed="true"] { background:var(--brand);border-color:var(--brand);color:#fff; }
  .btn { border:1px solid var(--line);background:var(--card);color:var(--ink2);font:inherit;font-weight:600;padding:9px 14px;border-radius:10px;cursor:pointer;white-space:nowrap; }
  .btn:hover:not(:disabled) { background:var(--bg); } .btn:disabled { opacity:.45;cursor:default; }
  .btn.primary { background:linear-gradient(135deg,#1E4D8C 0%,#2F6FD0 100%);border:none;color:#fff;flex:1;font-weight:700;padding:11px 16px;border-radius:12px;
    display:inline-flex;align-items:center;justify-content:center;gap:9px;letter-spacing:.1px;
    box-shadow:0 1px 2px rgba(16,42,92,.22),0 6px 14px -6px rgba(30,77,140,.55);transition:transform .12s ease,box-shadow .12s ease,filter .12s ease; }
  .btn.primary:hover:not(:disabled) { background:linear-gradient(135deg,#1E4D8C 0%,#2F6FD0 100%);filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 2px 4px rgba(16,42,92,.22),0 10px 20px -8px rgba(30,77,140,.6); }
  .btn.primary:active:not(:disabled) { transform:translateY(0);filter:brightness(.97);box-shadow:0 1px 2px rgba(16,42,92,.25); }
  .btn.primary:disabled { background:var(--line);color:var(--faint);box-shadow:none;opacity:1; }
  .btn:focus-visible { outline:2px solid #3B82F6;outline-offset:2px; }
  .ico-play::before { content:"";border-left:9px solid currentColor;border-top:5.5px solid transparent;border-bottom:5.5px solid transparent;margin-left:2px; }
  .ico-pause::before { content:"";width:3px;height:11px;background:currentColor;border-radius:1px;box-shadow:5px 0 0 currentColor;margin-right:5px; }
  .ico-pause { display:inline-flex;align-items:center;gap:6px; }
  .btn.mini { padding:4px 10px;font-size:12px; }
  input[type=text], input[type=number], textarea { width:100%;font:inherit;color:var(--ink);background:var(--card);border:1px solid var(--line);border-radius:9px;padding:7px 9px;outline:none; }
  input[type=number] { width:76px; }
  input[type=checkbox] { margin:0;width:15px;height:15px;flex:none;accent-color:var(--brand);cursor:pointer; }
  .coche { cursor:pointer; }
  textarea { resize:vertical;min-height:70px; }
  input:focus, textarea:focus { border-color:var(--brand); }
  .editeur { min-height:150px;max-height:320px;overflow:auto;background:var(--card);border:1px solid var(--line);border-radius:9px;padding:8px 10px;outline:none;color:var(--ink);word-break:break-word;resize:vertical; }
  .editeur:focus { border-color:var(--brand); }
  .editeur:empty::before { content:attr(data-placeholder);color:var(--faint); }
  .editeur a { color:var(--brand);text-decoration:underline; }
  .editeur p { margin:0; }
  .editeur[contenteditable="false"] { opacity:.75; }
  .gap { margin-top:8px; }
  .bar { height:6px;background:var(--track);border-radius:3px;margin-top:12px;overflow:hidden; }
  .bar > div { height:100%;width:0;background:var(--brand);border-radius:3px;transition:width .2s; }
  .notice { border-radius:10px;padding:9px 11px;background:var(--info-bg);color:var(--brand);margin-bottom:10px;font-size:12px; }
  .notice.ok { background:var(--ok-bg);color:var(--ok); } .notice.warn { background:var(--warn-bg);color:var(--warn); }
  .notice:empty { display:none; }
  .fichier { background:var(--card);border:1px solid var(--line);border-radius:9px;padding:7px 9px; }
  .results { max-height:220px;overflow-y:auto; } .res { display:flex;gap:10px;align-items:center;margin-bottom:7px; }
  .badge { min-width:28px;height:24px;padding:0 7px;border-radius:7px;display:grid;place-items:center;font-weight:700;font-size:11px;flex:none;background:var(--track);color:var(--muted); }
  .badge.envoye { background:var(--ok-bg);color:var(--ok); } .badge.verifier { background:var(--warn-bg);color:var(--warn); }
  .badge.envoi { background:var(--brand-soft);color:var(--brand); }
  .badge.deja { color:var(--faint); }
  .res .m { font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .res .n { font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .res > div:last-child { min-width:0;flex:1; } .empty { color:var(--faint);text-align:center;padding:14px 0; }
  .compact .q { font-size:24px;font-weight:700;color:var(--brand); }
  .compact .mail { font-weight:600;margin:6px 0 2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .compact .etape { color:var(--muted);font-size:12px;min-height:16px; }
  .compact .touch { color:#B45309;font-size:12px; }
  [hidden] { display:none !important; }
  `;

  var HTML = `
  <div class="panel">
    <div class="head" id="pp-head">
      <div class="logo">SM</div>
      <div style="flex:1;min-width:0"><h1>Smart Mail<span class="version" id="pp-version"></span></h1><div class="sub">Envoyer, comme à la main dans cet onglet : rien ne passe par un serveur</div></div>
      <button class="iconbtn" id="pp-gear" title="Paramètres">⚙</button>
      <button class="iconbtn" id="pp-min" title="Réduire">—</button>
    </div>
    <div class="perime" id="pp-perime" hidden>Extension mise à jour : rechargez cette page (Cmd+R ou F5) pour utiliser la nouvelle version.</div>
    <div class="body" id="pp-body">
      <div id="pp-prep">
        <div class="card spread" id="pp-reprise" hidden style="border-color:var(--brand)">
          <div style="min-width:0"><b id="pp-reprise-titre">Envoi en cours de route</b><div class="small" id="pp-reprise-texte" style="margin-top:2px"></div></div>
          <button class="btn" id="pp-abandon">Nouvel envoi</button>
        </div>
        <div class="card">
          <div class="spread"><p class="label">Destinataires</p><span class="small" id="pp-nb">0 adresse</span></div>
          <div id="pp-saisie">
            <label class="spread fichier gap" id="pp-annuaire" hidden style="margin-bottom:8px;cursor:pointer">
              <span class="row gap" style="min-width:0"><input type="checkbox" id="pp-annuaire-actif" checked style="margin:0"><span class="small" id="pp-annuaire-texte" style="min-width:0"></span></span>
              <span class="small" id="pp-annuaire-direct" style="color:var(--ok);font-weight:600;white-space:nowrap;flex:none">● en direct</span>
            </label>
            <textarea id="pp-dest" rows="3" placeholder="Collez des adresses : une par ligne, ou séparées par des virgules ou des points-virgules"></textarea>
            <div class="spread fichier gap" id="pp-fichier" hidden>
              <span class="small" id="pp-fichier-texte" style="min-width:0"></span>
              <button class="btn mini" id="pp-fichier-retirer">Retirer</button>
            </div>
            <div class="row gap"><button class="btn" id="pp-import">Importer un fichier texte</button><span class="small" id="pp-dest-info"></span></div>
            <input type="file" id="pp-file" accept=".txt,.csv,text/plain,text/csv" hidden>
            <p class="help">Fichier .txt avec une adresse par ligne (un .csv marche aussi). Gros fichiers acceptés, jusqu'à 200 Mo. Les doublons sont retirés.</p>
          </div>
          <div class="small" id="pp-liste-campagne" hidden></div>
        </div>
        <div class="card" id="pp-deja-carte">
          <div class="spread"><p class="label">Déjà envoyés</p><span class="small" id="pp-deja-nb">0 adresse</span></div>
          <div class="row" style="flex-wrap:wrap"><button class="btn mini" id="pp-deja-import">Importer un fichier</button><button class="btn mini" id="pp-deja-export">Télécharger</button><button class="btn mini" id="pp-deja-vider">Vider</button></div>
          <input type="file" id="pp-deja-file" accept=".txt,.csv,text/plain,text/csv" hidden>
          <p class="help">Ces personnes ne recevront plus rien : elles sont retirées de chaque nouvel envoi. La liste se remplit toute seule et le fichier publipostage-deja-envoyes.txt (Téléchargements) est mis à jour à chaque arrêt de l'envoi.</p>
        </div>
        <div class="card">
          <p class="label">Message</p>
          <input type="text" id="pp-objet" placeholder="Objet">
          <div id="pp-texte" class="editeur gap" contenteditable="true" role="textbox" aria-multiline="true" aria-label="Texte du message"
            data-placeholder="Écrivez ou collez votre mail ici : les liens restent cliquables"></div>
          <div class="row gap"><button class="btn mini" id="pp-lien" title="Sélectionnez un mot du message (ou cliquez dans un lien pour changer son adresse), puis cliquez">Ajouter ou changer un lien</button><button class="btn mini" id="pp-modele" title="Remet l'objet et le texte du message type">Message type</button></div>
          <div class="small" id="pp-nb-liens" style="margin-top:4px;overflow-wrap:anywhere"></div>
          <label class="row small gap coche"><input type="checkbox" id="pp-sans-signature"> Retirer le texte de la signature (nom, classe, ville) : le logo et les réseaux restent</label>
          <label class="row small gap coche"><input type="checkbox" id="pp-importance"> Importance élevée</label>
          <p class="help">Collez votre mail tel quel : liens, gras et listes sont conservés, et les adresses web (https://…, www.…) deviennent des liens. Le message type est déjà en place : modifiez-le librement, « Message type » le remet. Dans Outlook, l'adresse, l'objet et le message sont collés d'un coup, comme un copier-coller.</p>
        </div>
        <div class="card" id="pp-rythme">
          <p class="label">Rythme</p>
          <div class="row small">Pause entre deux mails : <input type="number" id="pp-pmin" min="0" step="1"> à <input type="number" id="pp-pmax" min="0" step="1"> s</div>
          <div class="row small gap">Maximum sur 24 h : <input type="number" id="pp-max" min="1" step="1"> <span id="pp-max-unite">mails</span></div>
          <div class="row small gap"><label class="row coche"><input type="checkbox" id="pp-cci"> Envoyer en Cci :</label><input type="number" id="pp-paquet" min="1" max="500" step="1"> personnes par mail</div>
          <p class="help" id="pp-estim"></p>
          <p class="help">L'envoi continue en arrière-plan (autre onglet, autre application, écran éteint) tant que Chrome et l'onglet Outlook restent ouverts. Le Mac ne se met pas en veille pendant l'envoi, sauf capot fermé sur batterie.</p>
        </div>
        <div class="row" id="pp-go-ligne" style="margin-bottom:10px"><button class="btn primary" id="pp-go">Lancer l'envoi</button></div>
        <div class="notice" id="pp-notice"></div>
      </div>
      <div id="pp-set" hidden>
        <div class="card">
          <p class="label">Si ça ne marche pas</p>
          <p class="help" style="margin-top:0">Le diagnostic ouvre un nouveau message, relève les champs de votre Outlook sans rien taper ni envoyer, referme le brouillon et télécharge un fichier à me renvoyer.</p>
          <button class="btn gap" id="pp-diag">Lancer le diagnostic</button>
        </div>
        <div class="card">
          <p class="label">Données</p>
          <p class="help" style="margin-top:0">Tout reste dans ce navigateur (stockage de l'extension) : campagne, suivi, réglages.</p>
          <button class="btn gap" id="pp-retour">Retour</button>
        </div>
        <div class="notice" id="pp-notice2"></div>
      </div>
      <div id="pp-live" class="compact" hidden>
        <div class="spread">
          <div class="q" id="pp-live-n">0 / 0</div>
          <div style="text-align:right"><div style="font-weight:600" id="pp-live-titre">Envoi en cours</div><div class="small" id="pp-live-eta"></div></div>
        </div>
        <div class="mail" id="pp-live-mail"></div>
        <div class="etape" id="pp-live-etape"></div>
        <div class="bar"><div id="pp-live-fill"></div></div>
        <div class="spread" style="margin-top:12px">
          <span class="touch">Laissez cet onglet au premier plan</span>
          <button class="btn" id="pp-pause">Pause</button>
        </div>
      </div>
      <div class="card" id="pp-suivi" style="padding:10px 12px;margin-top:10px">
        <div class="spread" style="margin-bottom:8px"><b>Suivi</b><span class="small" id="pp-suivi-resume"></span></div>
        <div class="results" id="pp-rows"><div class="empty">Les envois apparaîtront ici.</div></div>
        <div class="row gap"><button class="btn" id="pp-export" disabled>Exporter le suivi</button></div>
      </div>
    </div>
  </div>`;

  // ------------------------------------------------------------ montage
  var hote = document.createElement("div");
  hote.setAttribute("data-pp-panneau", "1");
  hote.style.cssText = "position:fixed;top:12px;right:12px;z-index:2147483647;";
  var ombre = hote.attachShadow({ mode: "open" });
  var st = document.createElement("style");
  st.textContent = CSS;
  ombre.appendChild(st);
  var analyse = new DOMParser().parseFromString(HTML, "text/html");
  while (analyse.body.firstChild) ombre.appendChild(analyse.body.firstChild);
  function monter() { if (!document.documentElement.contains(hote)) document.documentElement.appendChild(hote); }
  monter();
  try { new MutationObserver(monter).observe(document.documentElement, { childList: true }); } catch (e) {}
  // Le panneau « Adresses » occupe déjà le coin haut droit : on se place à sa gauche.
  function eviterAdresses() {
    if (hote.style.left || root.MEGA) return; // déjà déplacé à la main, ou onglets Smart Mail
    var ao = document.querySelector('[data-testid="ao-panel"]');
    if (ao) hote.style.right = Math.round(ao.getBoundingClientRect().width + 24) + "px";
  }
  eviterAdresses();
  setTimeout(eviterAdresses, 1500);
  setTimeout(eviterAdresses, 5000);

  var $ = function (id) { return ombre.getElementById(id); };
  var moteur = new Moteur(null);
  var reglages = Object.assign({}, Moteur.DEFAUTS);
  var vue = "prep";
  // Adresses du fichier importé, gardées à part (pas dans la zone de texte).
  var importe = { nom: "", adresses: [], vues: new Set(), doublons: 0 };
  var nbDestinataires = 0; // adresses à servir (hors « Déjà envoyés »), avant la création de l'envoi

  function afficher(v) {
    vue = v;
    $("pp-prep").hidden = v !== "prep";
    $("pp-set").hidden = v !== "set";
    $("pp-live").hidden = v !== "live";
    // La carte « Rythme » reste visible pendant l'envoi (sous la progression) :
    // un changement de pause ou de limite s'applique tout de suite.
    var rythme = $("pp-rythme");
    if (v === "live") $("pp-live").after(rythme);
    else $("pp-go-ligne").before(rythme);
    rythme.style.marginTop = v === "live" ? "12px" : "";
  }
  function notice(texte, type, id) { var n = $(id || "pp-notice"); n.textContent = texte || ""; n.className = "notice" + (type ? " " + type : ""); }
  function nb(n) { return Number(n).toLocaleString("fr-FR"); }
  function pluriel(n, un, plusieurs) { return nb(n) + " " + (n > 1 ? plusieurs : un); }
  // Valeur saisie ; si elle est vide ou invalide, la valeur en cours est gardée
  // (jamais une valeur par défaut substituée en silence).
  function nombre(id, actuel, min) { var v = parseInt($(id).value, 10); return isFinite(v) && v >= min ? v : actuel; }

  function lireReglages() {
    reglages.pauseMin = nombre("pp-pmin", reglages.pauseMin, 0);
    reglages.pauseMax = Math.max(reglages.pauseMin, nombre("pp-pmax", reglages.pauseMax, 0));
    reglages.maxJour = nombre("pp-max", reglages.maxJour, 1);
    reglages.paquet = Math.min(Moteur.PAQUET_MAX, nombre("pp-paquet", reglages.paquet, 1));
    reglages.cci = $("pp-cci").checked;
    reglages.importance = $("pp-importance").checked;
    reglages.sansTexteSignature = $("pp-sans-signature").checked;
    $("pp-max-unite").textContent = reglages.cci ? "destinataires" : "mails";
    return reglages;
  }
  // Affiche dans la case qu'on vient de quitter la valeur réellement utilisée
  // (jamais dans une autre case : on pourrait être en train d'y taper).
  function afficherReglageUtilise(id) {
    $(id).value = { "pp-pmin": reglages.pauseMin, "pp-pmax": reglages.pauseMax, "pp-max": reglages.maxJour, "pp-paquet": reglages.paquet }[id];
  }
  function poserReglages(r) {
    reglages = Object.assign({}, Moteur.DEFAUTS, r || {});
    $("pp-pmin").value = reglages.pauseMin; $("pp-pmax").value = reglages.pauseMax; $("pp-max").value = reglages.maxJour;
    $("pp-paquet").value = reglages.paquet; $("pp-cci").checked = !!reglages.cci;
    $("pp-importance").checked = !!reglages.importance; $("pp-sans-signature").checked = !!reglages.sansTexteSignature;
    $("pp-max-unite").textContent = reglages.cci ? "destinataires" : "mails";
  }

  // ------------------------------------------------------------ message
  // Message envoyé : nettoyé, adresses web en clair transformées en liens.
  function htmlMessage() { return M.lienifier($("pp-texte").innerHTML); }
  function texteMessage() { return M.texteBrut($("pp-texte").innerHTML); }
  // Nombre de liens et leurs adresses (sans doublon) : un ancien lien resté dans le message se voit.
  function majLiens() {
    var liens = M.liens(htmlMessage());
    var adresses = liens.map(function (l) { return l.href.replace(/^(https?:\/\/|mailto:)/i, ""); })
      .filter(function (a, i, t) { return t.indexOf(a) === i; });
    $("pp-nb-liens").textContent = liens.length ? pluriel(liens.length, "lien cliquable", "liens cliquables") + " → " + adresses.join(", ") : "";
  }
  // Remplace le contenu de l'éditeur par sa version avec liens (si elle diffère).
  function lienifierEditeur() {
    var ed = $("pp-texte");
    if (ed.getAttribute("contenteditable") !== "true") return;
    var avec = htmlMessage();
    if (avec !== M.nettoyer(ed.innerHTML)) { ed.innerHTML = avec; sauverBrouillon(); }
    majLiens();
  }
  // Insère du HTML (déjà nettoyé) au curseur de l'éditeur.
  function insererDansEditeur(html) {
    var ed = $("pp-texte");
    var sel = ombre.getSelection ? ombre.getSelection() : root.getSelection();
    var r = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    if (!r || !ed.contains(r.commonAncestorContainer)) { r = document.createRange(); r.selectNodeContents(ed); r.collapse(false); }
    r.deleteContents();
    var frag = r.createContextualFragment(html);
    var dernier = frag.lastChild;
    r.insertNode(frag);
    if (dernier && sel) { var r2 = document.createRange(); r2.setStartAfter(dernier); r2.collapse(true); sel.removeAllRanges(); sel.addRange(r2); }
  }
  function collerDansEditeur(donnees) {
    var html = donnees.getData("text/html");
    var propre = M.lienifier(html ? html : M.texteEnHtml(donnees.getData("text/plain")));
    insererDansEditeur(propre);
    estimer(); majLiens(); sauverBrouillon();
  }
  $("pp-texte").addEventListener("paste", function (e) {
    if (!e.clipboardData) return;
    e.preventDefault();
    collerDansEditeur(e.clipboardData);
  });
  $("pp-texte").addEventListener("drop", function (e) {
    if (!e.dataTransfer) return;
    e.preventDefault();
    collerDansEditeur(e.dataTransfer);
  });
  $("pp-texte").addEventListener("input", function () { estimer(); majLiens(); sauverBrouillon(); });
  $("pp-texte").addEventListener("blur", lienifierEditeur);

  // « Ajouter ou changer un lien » : le texte sélectionné dans le message devient un
  // lien. Sélection (ou simple curseur) dans un lien existant : c'est ce lien, en
  // entier, qui change d'adresse. Sinon l'ancien lien restait autour du nouveau,
  // vide et invisible, et l'envoi s'arrêtait (« Lien absent du mail »).
  var selectionLien = null;
  function lienAutour(n) {
    var el = n && (n.nodeType === 1 ? n : n.parentElement);
    var a = el && el.closest("a");
    return a && $("pp-texte").contains(a) ? a : null;
  }
  function selectionDansEditeur() {
    var sel = ombre.getSelection ? ombre.getSelection() : root.getSelection();
    if (!sel || !sel.rangeCount) return null;
    var r = sel.getRangeAt(0);
    return $("pp-texte").contains(r.commonAncestorContainer) ? r.cloneRange() : null;
  }
  // Le clic sur le bouton fait perdre la sélection : on la note juste avant.
  $("pp-lien").addEventListener("mousedown", function (e) { e.preventDefault(); selectionLien = selectionDansEditeur(); });
  $("pp-lien").onclick = function () {
    if ($("pp-texte").getAttribute("contenteditable") !== "true") return;
    var r = selectionLien || selectionDansEditeur();
    selectionLien = null;
    var debut = r && lienAutour(r.startContainer), fin = r && lienAutour(r.endContainer), ancien = debut || fin;
    if (!r || (r.collapsed && !ancien)) { notice("Sélectionnez d'abord, dans le message, le mot ou la phrase à transformer en lien (ou cliquez dans un lien pour changer son adresse).", "warn"); return; }
    var adresse = (root.prompt("Adresse du lien (https://…, www.… ou mailto:…) :", ancien ? ancien.getAttribute("href") : "https://") || "").trim();
    if (!adresse || adresse === "https://") return;
    if (/^www\./i.test(adresse)) adresse = "https://" + adresse;
    if (!/^(https?:\/\/[^\s]+|mailto:[^\s]+@[^\s]+)$/i.test(adresse)) { notice("Adresse de lien non valide : " + adresse, "warn"); return; }
    if (debut) r.setStartBefore(debut);
    if (fin) r.setEndAfter(fin);
    var contenu = r.extractContents();
    [].slice.call(contenu.querySelectorAll("a")).forEach(function (x) {
      while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
      x.remove();
    });
    var a = document.createElement("a");
    a.setAttribute("href", adresse);
    a.appendChild(contenu);
    r.insertNode(a);
    $("pp-texte").innerHTML = M.nettoyer($("pp-texte").innerHTML);
    estimer(); majLiens(); sauverBrouillon();
    notice(ancien ? "Lien changé : " + adresse : "Lien ajouté.", "ok");
  };

  // « Message type » (modele.js) : objet et texte remis, à modifier ensuite.
  // Pendant un envoi, seul le texte change (l'objet est verrouillé).
  function poserModele() {
    if (!$("pp-objet").disabled) $("pp-objet").value = Modele.objet;
    $("pp-texte").innerHTML = M.nettoyer(Modele.html);
  }
  $("pp-modele").onclick = function () {
    if ($("pp-texte").getAttribute("contenteditable") !== "true") return;
    if (texteMessage().trim() && htmlMessage() !== M.lienifier(Modele.html) && !root.confirm("Remplacer le message par le message type ? Vos modifications seront perdues.")) return;
    poserModele();
    estimer(); majLiens(); sauverBrouillon();
    notice("Message type remis : vous pouvez le modifier.", "ok");
  };

  // ------------------------------------------------------------ destinataires
  // Smart Mail : annuaire commun, rempli en direct par le relevé d'adresses
  // (mega.js appelle PP.Panneau.majAnnuaire). Utilisé si la case est cochée.
  var annuaire = { adresses: [], vues: new Set(), actif: true, present: false };
  // Adresses de l'annuaire qui ne sont pas déjà dans le fichier importé.
  function deAnnuaire() {
    if (!annuaire.actif) return [];
    return annuaire.adresses.filter(function (a) { return !importe.vues.has(a.toLowerCase()); });
  }
  // Adresses saisies dans la zone de texte qui ne sont pas déjà dans le fichier importé ni l'annuaire.
  function saisies() {
    var r = Moteur.extraireAdresses($("pp-dest").value, new Set());
    var nouvelles = r.adresses.filter(function (a) { var k = a.toLowerCase(); return !importe.vues.has(k) && !(annuaire.actif && annuaire.vues.has(k)); });
    return { adresses: nouvelles, doublons: r.doublons + (r.adresses.length - nouvelles.length), invalides: r.invalides || 0 };
  }
  // Adresses de l'envoi : saisies, importées puis de l'annuaire, sans celles de « Déjà envoyés ».
  function listeFinale() { return saisies().adresses.concat(importe.adresses, deAnnuaire()).filter(function (a) { return !moteur.estDeja(a); }); }

  function majAnnuaireTexte() {
    $("pp-annuaire").hidden = !annuaire.present;
    $("pp-annuaire-actif").checked = annuaire.actif;
    $("pp-annuaire-texte").textContent = "Annuaire commun : " + pluriel(annuaire.adresses.length, "personne relevée", "personnes relevées");
  }

  function majDestinataires() {
    var s = saisies();
    var tous = s.adresses.concat(importe.adresses, deAnnuaire()), deja = 0;
    for (var i = 0; i < tous.length; i++) if (moteur.estDeja(tous[i])) deja++;
    nbDestinataires = tous.length - deja;
    $("pp-nb").textContent = pluriel(nbDestinataires, "adresse", "adresses");
    var doublons = s.doublons + importe.doublons, infos = [];
    if (doublons) infos.push(pluriel(doublons, "doublon retiré", "doublons retirés"));
    if (deja) infos.push(pluriel(deja, "déjà envoyée retirée", "déjà envoyées retirées"));
    // Adresses mal formées (accents, caractères interdits) : JAMAIS un morceau envoyé à une autre personne.
    if (s.invalides) infos.push(pluriel(s.invalides, "adresse invalide ignorée (accents ou caractères interdits)", "adresses invalides ignorées (accents ou caractères interdits)"));
    $("pp-dest-info").textContent = infos.join(", ");
    $("pp-fichier").hidden = !importe.nom;
    $("pp-fichier-texte").textContent = importe.nom ? importe.nom + " : " + pluriel(importe.adresses.length, "adresse", "adresses") : "";
    majAnnuaireTexte();
    estimer();
  }

  // Estimation : environ 12 s de gestes (ouvrir, coller, envoyer ; 15 s en Cci) + pause
  // moyenne, par mail ; jours selon la limite (en destinataires pour la Cci).
  function estimer() {
    var cci = !!reglages.cci, paquet = cci ? Moteur.taillePaquet(reglages) : 1, gestes = cci ? 15 : 12;
    var unite = cci ? " destinataires" : " mails";
    var n = moteur.campagne ? Moteur.compter(moteur.campagne).attente : nbDestinataires;
    var parMail = gestes + (reglages.pauseMin + reglages.pauseMax) / 2;
    var txt = "Environ " + Math.round(parMail) + " s par mail (pause comprise)";
    if (n) {
      var mails = Math.ceil(n / paquet);
      var parJour = Math.min(reglages.maxJour, Math.floor((24 * 3600) / parMail) * paquet);
      txt += ", soit " + pluriel(n, "destinataire", "destinataires") + (cci ? " (" + pluriel(mails, "mail", "mails") + " en Cci)" : "") + " en " +
        (n <= parJour ? Math.max(1, Math.round((mails * parMail) / 60)) + " min" : "environ " + pluriel(Math.ceil(n / parJour), "jour", "jours") + " à " + nb(parJour) + unite + " par 24 h");
    }
    // Ce qui sera réellement appliqué (après correction d'un maximum plus petit que le minimum).
    txt += ". Pause appliquée : " + (reglages.pauseMin === reglages.pauseMax ? reglages.pauseMin + " s" : "entre " + reglages.pauseMin + " et " + reglages.pauseMax + " s") +
      " après chaque mail (en plus d'environ " + gestes + " s pour ouvrir, coller et envoyer), au plus " + nb(reglages.maxJour) + unite + " sur 24 h" +
      (cci ? ", " + pluriel(paquet, "personne", "personnes") + " en Cci par mail (« À » reste vide)" : "");
    $("pp-estim").textContent = txt + ".";
  }

  var minuterie = null;
  function sauverBrouillon() {
    clearTimeout(minuterie);
    minuterie = setTimeout(function () {
      try {
        var o = {};
        o[CLE_BROUILLON] = { dest: $("pp-dest").value, objet: $("pp-objet").value, html: $("pp-texte").innerHTML, reglages: lireReglages() };
        api.storage.local.set(o);
        // Édité pendant une campagne (hors envoi en cours) : le texte s'applique aux envois restants.
        if (moteur.campagne && !moteur.enCours) moteur.majMessage(htmlMessage());
      } catch (e) {}
    }, 300);
  }
  function sauverImport() {
    try { var o = {}; o[CLE_IMPORT] = { nom: importe.nom, adresses: importe.adresses, doublons: importe.doublons }; api.storage.local.set(o); } catch (e) {}
  }
  function poserImport(nom, adresses, doublons) {
    importe = { nom: nom || "", adresses: adresses || [], vues: new Set(), doublons: doublons || 0 };
    for (var i = 0; i < importe.adresses.length; i++) importe.vues.add(importe.adresses[i].toLowerCase());
  }

  // ------------------------------------------------------------ déjà envoyés
  function majDeja() {
    var n = moteur.dejaListe.length;
    $("pp-deja-nb").textContent = pluriel(n, "adresse", "adresses");
    $("pp-deja-export").disabled = !n;
    $("pp-deja-vider").disabled = !n;
  }

  // ------------------------------------------------------------ suivi
  function libelleStatut(s) { return { attente: "…", envoi: "→", envoye: "✓", verifier: "!", deja: "=" }[s] || s; }
  function heure(t) { return t ? new Date(t).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : ""; }
  function rendreSuivi() {
    var c = moteur.campagne, box = $("pp-rows");
    box.textContent = "";
    if (!c || !c.total) {
      var e = document.createElement("div"); e.className = "empty"; e.textContent = "Les envois apparaîtront ici."; box.appendChild(e);
      $("pp-suivi-resume").textContent = ""; $("pp-export").disabled = true; return;
    }
    var n = Moteur.compter(c);
    $("pp-suivi-resume").textContent = nb(n.envoye) + " envoyé(s) sur " + nb(n.total) + (n.verifier ? ", " + nb(n.verifier) + " à vérifier" : "") + (n.deja ? ", " + nb(n.deja) + " déjà servi(s) avant" : "");
    $("pp-export").disabled = false;
    var suivi = moteur.lignes(250, 15);
    suivi.lignes.forEach(function (d) {
      var r = document.createElement("div"); r.className = "res";
      var b = document.createElement("div"); b.className = "badge " + d.s; b.textContent = libelleStatut(d.s);
      var t = document.createElement("div");
      var m = document.createElement("div"); m.className = "m"; m.textContent = d.a; t.appendChild(m);
      var info = d.s === "envoye" ? "envoyé " + (d.c ? "en Cci " : "") + "le " + heure(d.t) : d.s === "verifier" ? "à vérifier : " + (d.d || "") :
        d.s === "deja" ? "déjà envoyé avant : sauté" : d.d ? "en attente : " + d.d : "";
      if (info) { var x = document.createElement("div"); x.className = "n"; x.textContent = info; x.title = info; t.appendChild(x); }
      r.appendChild(b); r.appendChild(t); box.appendChild(r);
    });
    if (suivi.reste) { var p = document.createElement("div"); p.className = "empty"; p.textContent = "et " + pluriel(suivi.reste, "autre destinataire", "autres destinataires") + " en attente"; box.appendChild(p); }
  }

  // Une campagne existe : on reprend, on ne modifie pas la liste. Le message, lui,
  // reste modifiable et s'applique aux destinataires encore en attente (synchro
  // dans sauverBrouillon via moteur.majMessage).
  function etatCampagne() {
    var c = moteur.campagne;
    var enRoute = !!(c && c.curseur < c.total);
    $("pp-reprise").hidden = !c;
    $("pp-saisie").hidden = !!c;
    $("pp-liste-campagne").hidden = !c;
    $("pp-objet").disabled = !!c;
    $("pp-texte").setAttribute("contenteditable", "true"); // modifiable même pendant une campagne
    if (c) {
      var n = Moteur.compter(c);
      $("pp-reprise-titre").textContent = enRoute ? "Envoi en cours de route" : "Envoi terminé";
      $("pp-reprise-texte").textContent = enRoute
        ? nb(n.envoye) + " envoyé(s), " + nb(n.attente) + " restant(s)" + (n.verifier ? ", " + nb(n.verifier) + " à vérifier" : "") + (n.deja ? ", " + nb(n.deja) + " déjà servi(s) avant" : "") + ". Objet : " + c.objet
        : "Terminé : " + nb(n.envoye) + " envoyé(s)" + (n.verifier ? ", " + nb(n.verifier) + " à vérifier" : "") + ".";
      $("pp-nb").textContent = pluriel(c.total, "adresse", "adresses");
      $("pp-liste-campagne").textContent = pluriel(c.total, "destinataire", "destinataires") + " dans cet envoi.";
      $("pp-objet").value = c.objet;
      $("pp-texte").innerHTML = M.nettoyer(c.html);
      poserReglages(c.reglages);
      majLiens();
    } else majDestinataires();
    $("pp-go").textContent = enRoute ? "Reprendre l'envoi" : "Lancer l'envoi";
    $("pp-go").disabled = !!c && !enRoute;
    $("pp-abandon").textContent = enRoute ? "Abandonner" : "Nouvel envoi";
    estimer();
    rendreSuivi();
  }

  // ------------------------------------------------------------ interface du moteur
  function progression(n) {
    $("pp-live-n").textContent = nb(n.envoye) + " / " + nb(n.total);
    $("pp-live-fill").style.width = (n.total ? Math.round((100 * (n.envoye + n.verifier)) / n.total) : 0) + "%";
  }
  moteur.ui = {
    debut: function () {
      afficher("live");
      $("pp-live-titre").textContent = "Envoi en cours";
      $("pp-pause").disabled = false;
      $("pp-live-eta").textContent = "";
    },
    courant: function (adresse, n) { $("pp-live-mail").textContent = adresse; $("pp-live-eta").textContent = ""; progression(n); },
    etape: function (texte) { $("pp-live-etape").textContent = texte; },
    pause: function (s, tiree, r) {
      $("pp-live-etape").textContent = "Pause de " + tiree + " s avant le mail suivant (réglage : " + r.pauseMin + " à " + r.pauseMax + " s)";
      $("pp-live-eta").textContent = "suivant dans " + s + " s";
      hote.dataset.ppPause = String(tiree); // repère pour les tests
    },
    etat: function (c) { progression(Moteur.compter(c)); rendreSuivi(); majDeja(); },
  };

  // ------------------------------------------------------------ actions
  ["pp-pmin", "pp-pmax", "pp-max", "pp-paquet"].forEach(function (id) {
    // Pris en compte tout de suite, même pendant un envoi (pause en cours comprise).
    $(id).addEventListener("input", function () { lireReglages(); estimer(); sauverBrouillon(); if (moteur.campagne) moteur.reglerRythme(reglages); });
    $(id).addEventListener("change", function () { lireReglages(); afficherReglageUtilise(id); estimer(); });
  });
  // Cases à cocher : appliquées au mail suivant, même pendant un envoi.
  ["pp-cci", "pp-importance", "pp-sans-signature"].forEach(function (id) {
    $(id).addEventListener("change", function () { lireReglages(); estimer(); sauverBrouillon(); if (moteur.campagne) moteur.reglerRythme(reglages); });
  });
  var minuterieDest = null;
  $("pp-dest").addEventListener("input", function () { clearTimeout(minuterieDest); minuterieDest = setTimeout(majDestinataires, 250); sauverBrouillon(); });
  ["pp-objet"].forEach(function (id) { $(id).addEventListener("input", function () { estimer(); sauverBrouillon(); }); });

  $("pp-import").onclick = function () { $("pp-file").click(); };
  $("pp-file").onchange = async function () {
    var f = $("pp-file").files && $("pp-file").files[0];
    $("pp-file").value = "";
    if (!f) return;
    if (f.size > TAILLE_MAX) { notice(f.name + " dépasse 200 Mo : découpez-le en plusieurs fichiers.", "warn"); return; }
    notice("Lecture de " + f.name + "…", "");
    var texte;
    try { texte = await f.text(); } catch (e) { notice("Lecture du fichier impossible.", "warn"); return; }
    var r = Moteur.extraireAdresses(texte, new Set());
    if (!r.adresses.length) { notice("Aucune adresse trouvée dans " + f.name + ".", "warn"); return; }
    poserImport(f.name, r.adresses, r.doublons);
    sauverImport();
    majDestinataires();
    notice(pluriel(r.adresses.length, "adresse importée", "adresses importées") + " depuis " + f.name + (r.doublons ? " (" + pluriel(r.doublons, "doublon retiré", "doublons retirés") + ")" : "") + ".", "ok");
  };
  $("pp-fichier-retirer").onclick = function () {
    poserImport("", [], 0);
    sauverImport();
    majDestinataires();
    notice("Fichier retiré de la liste.", "");
  };

  // Liste « Déjà envoyés » : importer un fichier (ajouté à la liste), télécharger, vider.
  $("pp-deja-import").onclick = function () { $("pp-deja-file").click(); };
  $("pp-deja-file").onchange = async function () {
    var f = $("pp-deja-file").files && $("pp-deja-file").files[0];
    $("pp-deja-file").value = "";
    if (!f) return;
    if (f.size > TAILLE_MAX) { notice(f.name + " dépasse 200 Mo.", "warn"); return; }
    var texte;
    try { texte = await f.text(); } catch (e) { notice("Lecture du fichier impossible.", "warn"); return; }
    var r = Moteur.extraireAdresses(texte, new Set());
    if (!r.adresses.length) { notice("Aucune adresse trouvée dans " + f.name + ".", "warn"); return; }
    var n = await moteur.ajouterDeja(r.adresses);
    majDeja();
    if (moteur.campagne) estimer(); else majDestinataires();
    var connues = r.adresses.length - n;
    notice(pluriel(n, "adresse ajoutée", "adresses ajoutées") + " à « Déjà envoyés » depuis " + f.name + (connues ? " (" + pluriel(connues, "était déjà dans la liste", "étaient déjà dans la liste") + ")" : "") + ".", "ok");
  };
  $("pp-deja-export").onclick = async function () {
    var r = await moteur.ecrireFichierDeja();
    notice(pluriel(moteur.dejaListe.length, "adresse", "adresses") + " dans " + r.nom + " (dossier Téléchargements).", "ok");
  };
  $("pp-deja-vider").onclick = async function () {
    if (moteur.enCours) return;
    var n = moteur.dejaListe.length;
    if (!root.confirm("Vider la liste « Déjà envoyés » (" + pluriel(n, "adresse", "adresses") + ") ? Ces personnes pourront de nouveau recevoir un mail. Le fichier déjà téléchargé n'est pas effacé.")) return;
    await moteur.viderDeja();
    majDeja();
    if (moteur.campagne) estimer(); else majDestinataires();
    notice("Liste « Déjà envoyés » vidée.", "");
  };

  function verifierMessage() {
    if (!$("pp-objet").value.trim()) return "Ajoutez un objet.";
    if (!texteMessage().trim()) return "Écrivez ou collez le texte du message.";
    return null;
  }

  async function executer(promesse) {
    $("pp-go").disabled = true;
    if (root.__smVerrou && root.__smVerrou !== "envoi") { $("pp-go").disabled = false; notice("Un relevé d'adresses est en cours : arrêtez-le avant d'envoyer.", "warn"); return; }
    root.__smVerrou = "envoi";
    var r;
    try { r = await promesse(); }
    catch (e) { r = { message: "Erreur : " + (e && e.message ? e.message : e), type: "warn" }; }
    finally { if (root.__smVerrou === "envoi") root.__smVerrou = null; }
    afficher("prep");
    etatCampagne();
    notice(r.message, r.type);
  }

  // Un relevé d'adresses en cours dans N'IMPORTE QUELLE fenêtre ou onglet (marqueur
  // « actif-outlook » / « actif-gmail » rafraîchi pendant le relevé) : pas d'envoi.
  async function releveAilleurs() {
    try {
      var o = await api.storage.local.get(["actif-outlook", "actif-gmail"]);
      return ["actif-outlook", "actif-gmail"].some(function (k) { return o[k] && Date.now() - o[k] < 30000; });
    } catch (e) { return false; }
  }
  $("pp-go").onclick = async function () {
    if (moteur.enCours) return;
    if (root.__smVerrou && root.__smVerrou !== "envoi") { notice("Un relevé d'adresses est en cours : arrêtez-le avant d'envoyer (il tape dans le champ « À »).", "warn"); return; }
    if ((root.AOPanneau && root.AOPanneau.enCours()) || await releveAilleurs()) { notice("Un relevé d'adresses est en cours : arrêtez-le avant d'envoyer (il tape dans le champ « À »).", "warn"); return; }
    if (moteur.enCours) return;
    if (!moteur.campagne) {
      var probleme = verifierMessage();
      var dest = listeFinale();
      if (!dest.length) probleme = saisies().adresses.length + importe.adresses.length + deAnnuaire().length ? "Toutes ces adresses ont déjà reçu le message (liste « Déjà envoyés »)." : "Ajoutez au moins une adresse.";
      if (probleme) { notice(probleme, "warn"); return; }
      if (await lireCle("pp-campagne")) { notice("Un envoi existe déjà (ouvert ou lancé dans un autre onglet) : rechargez cette page pour le voir, plutôt que d'en créer un second.", "warn"); return; }
      executer(async function () {
        await moteur.creer($("pp-objet").value.trim(), htmlMessage(), dest, lireReglages());
        etatCampagne();
        return moteur.lancer();
      });
      return;
    }
    executer(function () { return moteur.lancer(); });
  };
  $("pp-pause").onclick = function () {
    moteur.arret = true;
    $("pp-pause").disabled = true;
    $("pp-live-etape").textContent = "Pause demandée : arrêt après ce mail.";
  };
  $("pp-abandon").onclick = async function () {
    if (moteur.enCours) return;
    if (await moteur.envoiAilleurs()) { notice("Cet envoi tourne dans un autre onglet Outlook : mettez-le en pause là-bas avant de l'abandonner.", "warn"); return; }
    var c = moteur.campagne;
    if (c && c.curseur < c.total && !root.confirm("Abandonner cet envoi ? Les destinataires restants ne recevront rien.")) return;
    await moteur.abandonner();
    etatCampagne();
    notice("Prêt pour un nouvel envoi. Les adresses et le message restent dans les champs.", "");
  };

  function telecharger(nom, texte, type) {
    var url = URL.createObjectURL(new Blob([texte], { type: type }));
    var a = document.createElement("a");
    a.href = url; a.download = nom; a.style.display = "none";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
  }
  function date() { var d = new Date(), p = function (x) { return String(x).padStart(2, "0"); }; return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "-" + p(d.getHours()) + p(d.getMinutes()); }
  // CSV au séparateur « ; » (Excel français), cellules protégées contre les formules.
  function cellule(v) { var s = String(v == null ? "" : v); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; return '"' + s.replace(/"/g, '""') + '"'; }
  $("pp-export").onclick = function () {
    if (!moteur.campagne) return;
    var noms = { attente: "en attente", envoi: "en cours", envoye: "envoyé", verifier: "à vérifier", deja: "déjà envoyé avant" };
    var lignes = [["adresse", "statut", "date", "détail"].map(cellule).join(";")];
    moteur.toutes().forEach(function (d) { lignes.push([d.a, noms[d.s] || d.s, d.t ? new Date(d.t).toLocaleString("fr-FR") : "", d.d || (d.c ? "en Cci" : "")].map(cellule).join(";")); });
    telecharger("publipostage-suivi-" + date() + ".csv", "\uFEFF" + lignes.join("\r\n") + "\r\n", "text/csv;charset=utf-8");
  };

  $("pp-gear").onclick = function () { if (moteur.enCours) return; afficher(vue === "set" ? "prep" : "set"); };
  $("pp-retour").onclick = function () { afficher("prep"); };
  $("pp-diag").onclick = async function () {
    if (moteur.enCours) return;
    if ((root.AOPanneau && root.AOPanneau.enCours()) || await releveAilleurs()) { notice("Un relevé d'adresses est en cours : mettez-le en pause avant le diagnostic.", "warn", "pp-notice2"); return; }
    notice("Diagnostic en cours…", "", "pp-notice2");
    try {
      var rapport = await O.diagnostic();
      telecharger("publipostage-diagnostic-" + date() + ".json", JSON.stringify(rapport, null, 2), "application/json");
      notice("Diagnostic téléchargé (dossier Téléchargements). Envoyez-le-moi si l'envoi ne marche pas.", "ok", "pp-notice2");
    } catch (e) { notice("Diagnostic : " + (e && e.message ? e.message : e), "warn", "pp-notice2"); }
  };

  var replie = false;
  $("pp-min").onclick = function () { replie = !replie; $("pp-body").hidden = replie; $("pp-min").textContent = replie ? "+" : "—"; };

  // Déplacer le panneau en tirant l'en-tête.
  (function () {
    var tete = $("pp-head"), dx = 0, dy = 0, tire = false;
    tete.addEventListener("pointerdown", function (e) {
      if (e.target.closest("button")) return;
      tire = true; var r = hote.getBoundingClientRect(); dx = e.clientX - r.left; dy = e.clientY - r.top;
      tete.setPointerCapture(e.pointerId);
    });
    tete.addEventListener("pointermove", function (e) {
      if (!tire) return;
      hote.style.left = Math.max(0, Math.min(root.innerWidth - 80, e.clientX - dx)) + "px";
      hote.style.top = Math.max(0, Math.min(root.innerHeight - 40, e.clientY - dy)) + "px";
      hote.style.right = "auto";
    });
    tete.addEventListener("pointerup", function () { tire = false; });
  })();

  // Bouton de l'extension dans la barre du navigateur : afficher ou masquer le panneau.
  try {
    api.runtime.onMessage.addListener(function (m) {
      if (m && m.pp === "basculer" && !root.MEGA) hote.style.display = hote.style.display === "none" ? "" : "none";
    });
  } catch (e) {}

  // ------------------------------------------------------------ démarrage
  function lireCle(cle) { return new Promise(function (r) { try { api.storage.local.get(cle, function (o) { r(o && o[cle]); }); } catch (e) { r(null); } }); }
  (async function () {
    var b = await lireCle(CLE_BROUILLON);
    if (b) {
      $("pp-dest").value = b.dest || ""; $("pp-objet").value = b.objet || "";
      $("pp-texte").innerHTML = b.html ? M.nettoyer(b.html) : b.texte ? M.texteEnHtml(b.texte) : "";
    }
    // Objet ou texte vide : le message type, modifiable.
    if (!$("pp-objet").value.trim()) $("pp-objet").value = Modele.objet;
    if (!texteMessage().trim()) $("pp-texte").innerHTML = M.nettoyer(Modele.html);
    var imp = await lireCle(CLE_IMPORT);
    if (imp && Array.isArray(imp.adresses)) poserImport(imp.nom, imp.adresses, imp.doublons);
    poserReglages(b && b.reglages);
    majLiens();
    await moteur.charger();
    majDeja();
    etatCampagne();
    afficher("prep");
    hote.dataset.ppPret = "1"; // repère pour les tests
    // Smart Mail (mega.js) : l'annuaire commun, mis à jour en direct par le
    // relevé. Pendant un envoi en cours de route, la liste de l'envoi ne bouge
    // pas : les nouvelles personnes seront dans le prochain envoi.
    PP.Panneau = {
      hote: hote,
      enCours: function () { return moteur.enCours; },
      notice: function (texte, type) { notice(texte, type); },
      majAnnuaire: function (liste, actif) {
        annuaire.present = true;
        if (typeof actif === "boolean") annuaire.actif = actif;
        annuaire.adresses = liste.slice();
        annuaire.vues = new Set(liste.map(function (a) { return a.toLowerCase(); }));
        if (moteur.campagne) majAnnuaireTexte(); else majDestinataires();
      },
      surChoixAnnuaire: null // rappel de mega.js : case cochée ou décochée
    };
    $("pp-annuaire-actif").onchange = function () {
      annuaire.actif = $("pp-annuaire-actif").checked;
      majDestinataires();
      if (PP.Panneau.surChoixAnnuaire) PP.Panneau.surChoixAnnuaire(annuaire.actif);
    };
    try { document.dispatchEvent(new CustomEvent("mega-pret")); } catch (e) {}
  })();

  // Version affichée dans l'en-tête : elle change à chaque mise à jour de l'extension.
  try { $("pp-version").textContent = "v" + api.runtime.getManifest().version; } catch (e) {}
  // Extension rechargée (↻) alors que cette page tourne encore avec l'ancienne
  // version : on le dit, la page doit être rechargée.
  var veille = setInterval(function () {
    var vivante;
    try { vivante = !!(api.runtime && api.runtime.id); } catch (e) { vivante = false; }
    if (!vivante) { $("pp-perime").hidden = false; clearInterval(veille); }
  }, 3000);
})(typeof globalThis !== "undefined" ? globalThis : this);
