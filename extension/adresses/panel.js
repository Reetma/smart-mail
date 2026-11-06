// Panneau flottant injecté dans l'onglet de la messagerie (Shadow DOM, styles
// isolés). Reprend l'interface de la version Mac, sans ce qui ne sert pas sur
// le web : pas d'autorisation macOS, pas de choix de dossier, la messagerie est
// celle de l'onglet (détectée, pas choisie). Câble le même Moteur sur le même
// natif. Dernier fichier du script de contenu : tout le reste est déjà chargé.

(function (root) {
  "use strict";
  var api = root.browser || root.chrome;
  var Moteur = root.Moteur, natif = root.natif;
  var SOURCE = (root.AODom && root.AODom.SOURCE) || "outlook";
  // Nom du champ destinataire selon la messagerie (Gmail le nomme « Destinataires »).
  var CHAMP = SOURCE === "gmail" ? "« À » (Destinataires)" : "« À »";
  if (!Moteur || !natif) return;
  // Panneau + moteur uniquement dans le cadre haut. L'armement (armement.js)
  // tourne, lui, dans tous les cadres, au cas où la rédaction serait en iframe.
  if (window.top !== window.self) return;

  // Journal, verrou et « dernier fichier » propres à CETTE messagerie : Outlook et
  // Gmail ne partagent plus de clé (sinon l'un écraserait/supprimerait l'autre).
  var JOURNAL = "reprise-" + SOURCE, ACTIF = "actif-" + SOURCE, DERNIER = "dernier-" + SOURCE, BASE = "base-" + SOURCE;
  // Un seul passage par messagerie à la fois (même si plusieurs onglets) : un
  // marqueur rafraîchi pendant le passage, considéré périmé après 30 s sans signe
  // de vie. Un rechargement propre l'efface (pagehide), donc la reprise redémarre
  // tout de suite ; seul un vrai plantage laisse un marqueur qui s'efface en 30 s.
  var iOwnLock = false; // CET onglet tient-il le verrou ? (pour ne jamais libérer celui d'un autre)
  async function runActiveElsewhere() { try { var o = await api.storage.local.get(ACTIF); return !!o[ACTIF] && (Date.now() - o[ACTIF]) < 30000; } catch (e) { return false; } }
  function heartbeat() { try { iOwnLock = true; var o = {}; o[ACTIF] = Date.now(); api.storage.local.set(o); } catch (e) {} }
  function clearActive() { try { iOwnLock = false; api.storage.local.remove(ACTIF); } catch (e) {} }
  // Au déchargement, on ne libère le verrou QUE si c'est notre passage : sinon un
  // onglet inactif qui se ferme libérerait le verrou d'un autre onglet en cours.
  try { window.addEventListener("pagehide", function () { if (iOwnLock) clearActive(); }); } catch (e) {}

  var CSS = `
  :host { all: initial; }
  * { box-sizing: border-box; }
  .panel { --bg:#F6F8FC;--card:#FFF;--line:#E6EBF2;--ink:#0F172A;--ink2:#334155;--muted:#5B6A80;--faint:#94A3B8;
    --brand:#1E4D8C;--brand-dark:#173E73;--brand-soft:#E8F0FA;--info-bg:#EEF4FC;--ok-bg:#DCFCE7;--ok:#15803D;
    --warn-bg:#FEF3C7;--warn:#92400E;--chip:#EEF2F8;--chip-ink:#7A879A;--track:#E8EEF6;
    width:380px;max-width:92vw;max-height:88vh;overflow:auto;background:var(--card);color:var(--ink);
    border:1px solid var(--line);border-radius:14px;box-shadow:0 10px 40px rgba(15,23,42,.28);
    font:13px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; }
  @media (prefers-color-scheme: dark) {
    .panel { --bg:#0F141C;--card:#161D28;--line:#263042;--ink:#E7ECF3;--ink2:#C5CEDA;--muted:#93A1B5;--faint:#64748B;
      --brand:#4F86D1;--brand-dark:#3B72BE;--brand-soft:#1C2B42;--info-bg:#17253A;--ok-bg:#12301F;--ok:#5FD38D;
      --warn-bg:#33270F;--warn:#F2C46B;--chip:#1E2735;--chip-ink:#8391A6;--track:#222C3B; } }
  .head { display:flex;align-items:center;gap:10px;padding:12px 14px;border-bottom:1px solid var(--line);cursor:grab; }
  .logo { width:30px;height:30px;border-radius:8px;background:linear-gradient(45deg,#102A5C,#3B82F6);color:#fff;display:grid;place-items:center;font-weight:800;font-size:11.5px;letter-spacing:-.2px;flex:none;position:relative;font-family:'Avenir Next','Segoe UI',system-ui,sans-serif; }
  .logo::after { content:'✦';position:absolute;top:-1px;right:2px;font-size:8px;color:#FBBF24;line-height:1; }
  .head h1 { font-size:15px;margin:0;flex:1; } .head .sub { font-size:11px;color:var(--muted); }
  .iconbtn { border:0;background:transparent;color:var(--muted);font:inherit;font-size:16px;cursor:pointer;padding:2px 6px;border-radius:7px; }
  .head .iconbtn + .iconbtn { margin-left:-6px; }
  .iconbtn:hover { background:var(--bg); }
  .body { padding:14px; } .card { background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:10px; }
  .spread { display:flex;justify-content:space-between;align-items:center;gap:10px; }
  .row { display:flex;align-items:center;gap:8px; }
  .label { font-weight:600;color:var(--ink2);margin:0 0 5px; } .small { font-size:12px;color:var(--muted); }
  .help { font-size:12px;color:var(--muted);margin:5px 0 10px; }
  .pills { display:flex;flex-wrap:wrap;gap:6px; }
  .pill { border:1px solid var(--line);background:var(--card);color:var(--ink2);font:inherit;font-weight:600;padding:6px 12px;border-radius:999px;cursor:pointer; }
  .pill[aria-pressed="true"] { background:var(--brand);border-color:var(--brand);color:#fff; }
  .btn { border:1px solid var(--line);background:var(--card);color:var(--ink2);font:inherit;font-weight:600;padding:9px 14px;border-radius:10px;cursor:pointer;white-space:nowrap; }
  .btn:hover:not(:disabled) { background:var(--bg); } .btn:disabled { opacity:.45;cursor:default; }
  .btn.mini { padding:5px 10px;font-size:12px;border-radius:8px; }
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
  input[type=text] { width:100%;font:inherit;color:var(--ink);background:var(--card);border:1px solid var(--line);border-radius:9px;padding:7px 9px;outline:none; }
  input:focus { border-color:var(--brand); }
  .switch { display:flex;align-items:center;gap:9px;margin:7px 0;cursor:pointer;color:var(--ink2); }
  .switch input { width:34px;height:20px;-webkit-appearance:none;appearance:none;background:var(--track);border-radius:999px;position:relative;cursor:pointer;flex:none;margin:0; }
  .switch input::after { content:"";position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:#fff;transition:left .15s;box-shadow:0 1px 2px rgba(0,0,0,.25); }
  .switch input:checked { background:var(--brand); } .switch input:checked::after { left:16px; }
  .chips { display:grid;grid-template-columns:repeat(13,1fr);gap:4px;margin-top:10px; }
  .chip { height:24px;border-radius:6px;display:grid;place-items:center;font-weight:600;font-size:11px;background:var(--chip);color:var(--chip-ink);border:2px solid transparent; }
  .chip.now { background:var(--card);border-color:var(--brand);color:var(--brand); }
  .chip.ok { background:var(--brand);color:#fff; } .chip.none { background:var(--warn-bg);color:var(--warn); }
  .bar { height:6px;background:var(--track);border-radius:3px;margin-top:12px;overflow:hidden; }
  .bar > div { height:100%;width:0;background:var(--brand);border-radius:3px;transition:width .2s; }
  .notice { border-radius:10px;padding:9px 11px;background:var(--info-bg);color:var(--brand);margin-bottom:10px;font-size:12px; }
  .notice.ok { background:var(--ok-bg);color:var(--ok); } .notice.warn { background:var(--warn-bg);color:var(--warn); }
  .notice:empty { display:none; }
  .results { max-height:200px;overflow-y:auto; } .res { display:flex;gap:10px;margin-bottom:8px; }
  .badge { min-width:28px;height:28px;padding:0 6px;border-radius:8px;background:var(--brand-soft);color:var(--brand);display:grid;place-items:center;font-weight:700;font-size:11px;flex:none; }
  .res .m { font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .res .n { font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .res > div:last-child { min-width:0; } .empty { color:var(--faint);text-align:center;padding:14px 0; }
  .compact .live-tete { display:flex;align-items:center;gap:8px;margin-bottom:12px; }
  .compact .live-tete b { font-size:13.5px;color:var(--ink); }
  .compact .live-tete .duree { margin-left:auto;color:var(--muted);font-size:12px;font-variant-numeric:tabular-nums; }
  .compact .pouls { width:8px;height:8px;border-radius:50%;background:#16A34A;flex:none;position:relative; }
  .compact .pouls::after { content:"";position:absolute;inset:-4px;border-radius:50%;border:2px solid #16A34A;opacity:0;animation:ao-pouls 1.6s ease-out infinite; }
  @keyframes ao-pouls { 0% { transform:scale(.5);opacity:.7; } 100% { transform:scale(1.6);opacity:0; } }
  .compact .live-stats { display:grid;grid-template-columns:1.4fr 1fr;gap:8px; }
  .compact .tuile { background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:10px 12px; }
  .compact .tuile.principale { background:var(--brand-soft);border-color:transparent; }
  .compact .chiffre { font-size:26px;line-height:1.1;font-weight:800;color:var(--ink);font-variant-numeric:tabular-nums;letter-spacing:-.5px; }
  .compact .principale .chiffre { color:var(--brand);display:inline-block;transition:transform .18s ease; }
  .compact .chiffre.pop { transform:scale(1.12); }
  .compact .legende { font-size:11.5px;color:var(--muted);margin-top:2px; }
  .compact .live-piste { display:flex;align-items:center;gap:8px;margin:12px 0 0;min-width:0; }
  .compact .q { flex:none;max-width:45%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;background:var(--card);border:1px solid var(--line);
    color:var(--brand);font-weight:700;font-size:12px;letter-spacing:.4px;border-radius:999px;padding:2px 10px; }
  .compact .mail { color:var(--muted);font-size:12.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0; }
  .compact .pause-large { width:100%;margin-top:14px;display:inline-flex;align-items:center;justify-content:center;padding:10px 14px;border-radius:12px; }
  .compact .aide { margin:8px 0 0;text-align:center;font-size:11.5px;color:var(--faint); }
  [hidden] { display:none !important; }
  `;

  var HTML = `
  <div class="panel" part="panel">
    <div class="head" id="ao-head">
      <div class="logo">SM</div>
      <div style="flex:1"><h1 id="ao-title">Adresses</h1><div class="sub" id="ao-src"></div></div>
      <button class="iconbtn" id="ao-gear" title="Paramètres">⚙</button>
      <button class="iconbtn" id="ao-min" title="Réduire">—</button>
      <button class="iconbtn" id="ao-close" title="Fermer (revient au prochain chargement de la page)" aria-label="Fermer">✕</button>
    </div>
    <div class="body" id="ao-body">
      <!-- vue recherche -->
      <div id="ao-run-view">
        <div class="card" id="ao-arm-hint" style="padding:10px 12px">
          <b id="ao-arm-title">Prêt</b>
          <div class="small" id="ao-arm-msg" style="margin-top:3px">Ouvre un message et clique dans le champ ${CHAMP}. Je démarre tout seul après quelques secondes, ou clique « Démarrer ici » pour taper là tout de suite.</div>
        </div>
        <div hidden><span id="ao-count"></span><div id="ao-fill"></div></div>
        <div class="card spread" id="ao-resume" hidden style="border-color:var(--brand)">
          <div style="min-width:0"><b>En pause</b><div class="small" id="ao-resume-text" style="margin-top:2px"></div></div>
          <button class="btn" id="ao-fresh">Repartir de zéro</button>
        </div>
        <div class="row" style="margin-bottom:10px">
          <button class="btn primary ico-play" id="ao-run">Démarrer</button>
        </div>
        <div class="notice" id="ao-notice"></div>
        <div class="card" style="padding:10px 12px">
          <div class="spread" style="margin-bottom:8px"><b>Résultats</b><span class="small" id="ao-distinct"></span></div>
          <div class="results" id="ao-rows"><div class="empty" id="ao-empty">Les adresses apparaîtront ici.</div></div>
          <div class="spread gap" style="margin-top:10px;align-items:center">
            <span class="small">Tout reste dans l'extension (annuaire commun) : pas besoin d'exporter.</span>
            <span class="row" style="gap:6px;flex:none"><button class="btn mini" id="ao-export" disabled>Exporter</button><button class="btn mini" id="ao-copy" disabled>Copier</button></span>
          </div>
        </div>
      </div>
      <!-- vue paramètres : uniquement le diagnostic. Le reste est conservé mais
           masqué (le code interne s'en sert ; valeurs par défaut). -->
      <div id="ao-set-view" hidden>
        <div class="card">
          <p class="label">Historique et reprise</p>
          <p class="help" style="margin-top:0">« Importer » : un export d'adresses rejoint l'historique (liste connue qui grossit), ou un fichier de reprise recharge le passage pour continuer au même endroit. « Exporter pour continuer » enregistre le point de reprise du passage en cours.</p>
          <div class="row">
            <button class="btn" id="ao-import">Importer un fichier</button>
            <button class="btn" id="ao-save-resume">Exporter pour continuer</button>
          </div>
          <input type="file" id="ao-file" accept=".txt,.csv,.json,.jsonl" hidden>
          <p class="help" style="margin-bottom:6px">« Vider l'historique » oublie toutes les adresses connues (historique importé et derniers passages). Utile après une version qui lisait mal les adresses. Le passage interrompu n'est pas touché : « Repartir de zéro » s'en charge.</p>
          <button class="btn" id="ao-clear-base">Vider l'historique</button>
        </div>
        <div class="card">
          <p class="label">Si ça ne marche pas</p>
          <p class="help" style="margin-top:0">Le diagnostic vise le champ ${CHAMP} où est le curseur, tape « a », relève la structure de ta page et télécharge un fichier pour ajuster l'outil. Rien n'est envoyé. Tant que cette vue est ouverte, le passage ne démarre pas tout seul.</p>
          <button class="btn" id="ao-diag">Lancer le diagnostic</button>
        </div>
        <div hidden aria-hidden="true">
          <input type="text" id="ao-Prefix">
          <input type="text" id="ao-Keep">
          <input type="text" id="ao-Exclude">
          <label><input type="checkbox" id="ao-Dedupe"></label>
          <div class="pills" data-group="Speed">
            <button class="pill" data-value="fast">Rapide</button>
            <button class="pill" data-value="normal">Normal</button>
            <button class="pill" data-value="careful">Prudent</button>
          </div>
          <p id="ao-speed-help"></p>
          <div class="pills" data-group="Format">
            <button class="pill" data-value="txt">Texte</button>
            <button class="pill" data-value="csv">CSV</button>
            <button class="pill" data-value="both">Les deux</button>
          </div>
          <label><input type="checkbox" id="ao-CopyAtEnd"></label>
          <button class="btn" id="ao-reset">Réglages d'origine</button>
        </div>
      </div>
      <!-- vue compacte (pendant le passage) -->
      <div id="ao-compact-view" class="compact" hidden>
        <div class="live-tete"><span class="pouls"></span><b>Relevé en cours</b><span class="duree" id="ao-live-duree"></span></div>
        <div class="live-stats">
          <div class="tuile principale"><div class="chiffre" id="ao-live-n">0</div><div class="legende" id="ao-live-n-legende">personne trouvée</div></div>
          <div class="tuile"><div class="chiffre" id="ao-live-r">0</div><div class="legende" id="ao-live-r-legende">recherche</div></div>
        </div>
        <div class="live-piste" title="Recherche tapée dans « À », puis dernière adresse trouvée"><span class="q" id="ao-live-q">…</span><span class="mail" id="ao-live-mail"></span></div>
        <div hidden><span id="ao-live-count"></span><span id="ao-live-eta"></span><div id="ao-live-fill"></div></div>
        <button class="btn ico-pause pause-large" id="ao-stop">Pause</button>
        <p class="aide">Gardez cet onglet ouvert et visible pendant le relevé.</p>
      </div>
    </div>
  </div>`;

  // ------------------------------------------------------------ montage
  var host = document.createElement("div");
  host.setAttribute("data-testid", "ao-panel");
  host.style.cssText = "position:fixed;top:12px;right:12px;z-index:2147483647;";
  var shadow = host.attachShadow({ mode: "open" });
  // Construit sans innerHTML (contenu 100 % statique, mais on évite l'avertissement
  // de sécurité des linters) : une feuille de style + le balisage analysé.
  var styleEl = document.createElement("style"); styleEl.textContent = CSS; shadow.appendChild(styleEl);
  var parsed = new DOMParser().parseFromString(HTML, "text/html");
  while (parsed.body.firstChild) shadow.appendChild(parsed.body.firstChild);
  // Cliquer un bouton du panneau ne doit pas voler le focus du champ « À » armé.
  shadow.addEventListener("mousedown", function (e) { if (e.target && e.target.closest && e.target.closest("button")) e.preventDefault(); }, true);
  function mount() { if (!document.documentElement.contains(host)) document.documentElement.appendChild(host); }
  mount();
  // Les SPA (OWA/Gmail) remplacent parfois tout le corps : on se raccroche.
  var mountObs = null;
  try { mountObs = new MutationObserver(mount); mountObs.observe(document.documentElement, { childList: true }); } catch (e) {}

  var $ = function (id) { return shadow.getElementById(id); };
  var cfg, moteur, running = false, hasRun = false, resumable = false;
  var closed = false; // fermé par la croix : pour CETTE page seulement, rien n'est mémorisé
  var waitFn = (root.__AO_TEST && root.__AO_TEST.waitMs != null)
    ? function (ms) { return new Promise(function (r) { setTimeout(r, root.__AO_TEST.waitMs); }); } : null;

  function view(v) {
    $("ao-run-view").hidden = v !== "run";
    $("ao-set-view").hidden = v !== "set";
    $("ao-compact-view").hidden = v !== "compact";
  }
  function setPills(group, value) {
    shadow.querySelectorAll('[data-group="' + group + '"] .pill').forEach(function (b) {
      b.setAttribute("aria-pressed", String(b.dataset.value === String(value)));
    });
  }
  // Pendant un passage, les réglages sont verrouillés (sinon une modification en
  // cours fausserait le filtrage/dédoublonnage déjà appliqué aux lignes écrites).
  function setEditable(on) {
    ["Prefix", "Keep", "Exclude", "Dedupe", "CopyAtEnd"].forEach(function (k) { var el = $("ao-" + k); if (el) el.disabled = !on; });
    shadow.querySelectorAll(".pill").forEach(function (b) { b.disabled = !on; });
  }
  function applyConfig() {
    ["Speed", "Format"].forEach(function (g) { setPills(g, cfg[g]); });
    ["Prefix", "Keep", "Exclude"].forEach(function (k) { $("ao-" + k).value = cfg[k]; });
    ["Dedupe", "CopyAtEnd"].forEach(function (k) { $("ao-" + k).checked = !!cfg[k]; });
  }
  function readConfig() {
    cfg.Prefix = $("ao-Prefix").value.replace(/[;,]/g, "").trim();
    cfg.Keep = $("ao-Keep").value.trim();
    cfg.Exclude = $("ao-Exclude").value.trim();
    ["Dedupe", "CopyAtEnd"].forEach(function (k) { cfg[k] = $("ao-" + k).checked; });
  }
  var saveTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () { try { api.storage.local.set({ reglages: JSON.stringify(cfg) }); } catch (e) {} }, 300);
  }
  function update() {
    readConfig();
    var src = Moteur.sourceName(cfg);
    $("ao-title").textContent = "Smart Mail";
    $("ao-src").textContent = "Relever les adresses dans " + src + " — rien n'est envoyé";
    var sh = $("ao-speed-help"); if (sh) sh.textContent = Moteur.SPEEDS[cfg.Speed].text;
    if (!running) buildChips();
    save();
  }
  function buildChips() {
    var box = $("ao-chips"); // les pastilles A–Z ont été retirées : on ne garde que le compteur + la barre
    if (box) {
      box.innerHTML = "";
      for (var c = 0; c < 26; c++) {
        var d = document.createElement("div");
        d.className = "chip"; d.dataset.q = cfg.Prefix + String.fromCharCode(97 + c); d.textContent = String.fromCharCode(65 + c);
        box.appendChild(d);
      }
    }
    progress(0, 0, "");
  }
  function notice(text, kind) { var n = $("ao-notice"); n.textContent = text || ""; n.className = "notice" + (kind ? " " + kind : ""); }
  function progress(done, total, eta) {
    var label = Moteur.plural(done, "recherche", "recherches");
    $("ao-count").textContent = label; $("ao-live-count").textContent = label;
    var pct = total ? Math.round(100 * done / total) + "%" : "0%";
    $("ao-fill").style.width = pct; $("ao-live-fill").style.width = pct;
    $("ao-live-eta").textContent = eta || "";
    // Tuiles de la vue « Relevé en cours » : l'ETA du moteur vaut « durée · N adresses ».
    var parts = String(eta || "").split(" · "), n = parseInt(String(parts[1] || "").replace(/\D/g, ""), 10) || 0;
    $("ao-live-duree").textContent = parts[0] || "";
    $("ao-live-r").textContent = Number(done).toLocaleString("fr-FR");
    $("ao-live-r-legende").textContent = done > 1 ? "recherches" : "recherche";
    var nEl = $("ao-live-n");
    if (String(n) !== nEl.dataset.n) {
      if (n > (Number(nEl.dataset.n) || 0)) { nEl.classList.add("pop"); setTimeout(function () { nEl.classList.remove("pop"); }, 180); }
      nEl.dataset.n = String(n); nEl.textContent = n.toLocaleString("fr-FR");
      $("ao-live-n-legende").textContent = n > 1 ? "personnes trouvées" : "personne trouvée";
    }
  }
  function resetResults() {
    $("ao-rows").innerHTML = '<div class="empty" id="ao-empty">Les adresses apparaîtront ici.</div>';
    $("ao-distinct").textContent = "";
  }

  var ui = {
    debut: function () {
      running = true; hasRun = true; heartbeat();
      if (closed && moteur) moteur.stop = true; // fermé pendant l'amorçage : le passage s'arrête avant la 1re recherche
      try { if (root.AOArmement) root.AOArmement.suspend(); } catch (e) {} // la frappe du moteur ne doit pas ré-armer
      try { document.dispatchEvent(new Event("ao-run-begin")); } catch (e) {} // repère pour les fausses pages de test
      collapsed = false; $("ao-body").hidden = false; $("ao-min").textContent = "—"; // le panneau reste déplié (bouton Arrêter visible)
      resetResults(); buildChips(); setEditable(false);
      $("ao-live-q").textContent = "…"; $("ao-live-mail").textContent = "Préparation…"; $("ao-live-eta").textContent = "";
      $("ao-live-n").textContent = "0"; $("ao-live-n").dataset.n = "0"; $("ao-live-r").textContent = "0"; $("ao-live-duree").textContent = "";
      $("ao-live-n-legende").textContent = "personne trouvée"; $("ao-live-r-legende").textContent = "recherche";
      $("ao-stop").disabled = false;
      view("compact");
    },
    chip: function (q, state) { var d = shadow.querySelector('.chip[data-q="' + (window.CSS && window.CSS.escape ? window.CSS.escape(q) : q) + '"]'); if (d) d.className = "chip " + state; },
    live: function (q, mail) { if (q) $("ao-live-q").textContent = q; if (mail) $("ao-live-mail").textContent = mail; },
    row: function (q, mail, name) {
      try { if (root.MEGA) root.MEGA.trouve(mail); } catch (e) {} // Smart Mail : annuaire commun, en direct
      var e = $("ao-empty"); if (e) e.remove();
      var r = document.createElement("div"); r.className = "res";
      var b = document.createElement("div"); b.className = "badge"; b.textContent = (q.length > 10 ? q.slice(0, 9) + "…" : q).toUpperCase();
      var t = document.createElement("div");
      var m = document.createElement("div"); m.className = "m"; m.textContent = mail;
      var n = document.createElement("div"); n.className = "n"; n.textContent = name !== mail ? name : "";
      t.appendChild(m); if (n.textContent) t.appendChild(n);
      r.appendChild(b); r.appendChild(t);
      var box = $("ao-rows"); box.appendChild(r);
      // Affichage borné : on ne garde que les ~500 dernières lignes dans le DOM
      // (le fichier, l'export et le compteur « distinctes » gardent TOUT).
      while (box.childElementCount > 500) box.removeChild(box.firstElementChild);
    },
    progress: function (d, t, e) { progress(d, t, e); if (running) heartbeat(); },
    reglages: function (c) { cfg = c; applyConfig(); update(); },
    fin: function (msg, kind, info) {
      if (root.__smVerrou === "releve") root.__smVerrou = null;
      running = false; clearActive(); setEditable(true); view("run"); notice(msg, kind);
      try { if (root.AOArmement) { root.AOArmement.disarm(); if (!closed) root.AOArmement.resume(); } } catch (e) {} // prêt à ré-armer pour un prochain passage (sauf panneau fermé)
      if (info && info.distinct !== undefined) $("ao-distinct").textContent = Moteur.plural(info.distinct, "adresse distincte", "adresses distinctes");
      $("ao-run").disabled = false; $("ao-diag").disabled = false;
      try { api.storage.local.set({ reglages: JSON.stringify(cfg) }); } catch (e) {} // flush immédiat
      try { var dd = {}; dd[DERNIER] = cfg.Dernier || ""; api.storage.local.set(dd); } catch (e) {} // dernier fichier, par messagerie
      (async function () {
        try { if (moteur && moteur.distinct && moteur.distinct.length) await mergeIntoBase(moteur.distinct); } catch (e) {} // le passage s'ajoute à l'historique
        refreshResume();
      })();
    }
  };

  // ---- Historique (base-<SOURCE>) + import/reprise. L'« historique » montré à
  //      l'utilisateur (compteur, Export, Copier) = union de la base accumulée et
  //      des adresses du dernier passage. Rien dans le cœur n'est touché.
  var EMAIL = /(?<=^|[\s;,<>()\[\]"':=|])[A-Za-z0-9._%+'\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}(?=$|[\s;,<>()\[\]"':|.?!])/u; // trouve l'adresse ENTIÈRE dans une ligne
  function uniqSort(arr) { var u = Object.create(null); for (var i = 0; i < arr.length; i++) { var m = String(arr[i]).trim().toLowerCase(); if (m) u[m] = 1; } return Object.keys(u).sort(); }
  async function getBase() { try { var o = await api.storage.local.get(BASE); return o[BASE] ? String(o[BASE]).split(/\r?\n/).filter(Boolean) : []; } catch (e) { return []; } }
  async function mergeIntoBase(addrs) { var merged = uniqSort((await getBase()).concat(addrs)); try { var o = {}; o[BASE] = merged.join("\n"); await api.storage.local.set(o); } catch (e) {} return merged; }
  async function applyKnown() {
    var all = uniqSort((await getBase()).concat((moteur && moteur.distinct) || []));
    if (moteur) moteur.distinct = all;
    $("ao-export").disabled = all.length === 0; $("ao-copy").disabled = all.length === 0;
    $("ao-distinct").textContent = all.length ? Moteur.plural(all.length, "adresse distincte", "adresses distinctes") : "";
    return all;
  }
  // Un fichier importé : soit un POINT DE REPRISE (journal, 1re ligne JSON) qu'on
  // réinstalle pour « Reprendre » exactement, soit une LISTE d'adresses (.txt/.csv)
  // qu'on fusionne dans l'historique.
  async function importerTexte(text) {
    if (running) { notice("Attends la fin du passage pour importer.", "warn"); return; }
    var lines = String(text).replace(/^﻿/, "").split(/\r?\n/);
    var first = ""; for (var k = 0; k < lines.length; k++) { if (lines[k].trim()) { first = lines[k].trim(); break; } }
    var head = null; if (first.charAt(0) === "{") { try { head = JSON.parse(first); } catch (e) { head = null; } }
    if (head && (head.Base !== undefined || head.Started !== undefined)) {
      try { var oj = {}; oj[JOURNAL] = text; await api.storage.local.set(oj); } catch (e) {}
      try { var j = Moteur.parseJournal(text); if (j && j.keptList) await mergeIntoBase(j.keptList); } catch (e) {}
      await refreshResume();
      notice("Reprise chargée. Clique « Reprendre ici » pour continuer au même endroit.", "ok");
      return;
    }
    // On extrait l'adresse de CHAQUE ligne, quel que soit le format : adresse
    // seule, « Nom <adresse> » (l'export de l'application), « - Nom <adresse> »,
    // ou une colonne CSV. On prend la première adresse trouvée sur la ligne ;
    // l'en-tête CSV (sans adresse) est ignoré d'office.
    var addrs = [];
    for (var i = 0; i < lines.length; i++) {
      var m = lines[i].match(EMAIL);
      if (m) addrs.push(m[0].toLowerCase());
    }
    if (!addrs.length) { notice("Aucune adresse reconnue dans ce fichier.", "warn"); return; }
    var before = (await getBase()).length;
    var merged = await mergeIntoBase(addrs);
    await applyKnown();
    notice((merged.length - before) + " adresse(s) ajoutée(s) à l'historique (" + merged.length + " au total). Export et Copier les incluent.", "ok");
  }

  async function refreshResume() {
    var r = moteur ? await moteur.reprise() : null;
    if (moteur) await moteur.connues(cfg); // pose moteur.distinct depuis le journal/dernier fichier
    await applyKnown();                     // union avec l'historique importé/accumulé (base-<SOURCE>)
    if (!running) save();
    resumable = !!r;
    $("ao-resume").hidden = !r;
    if (r) $("ao-resume-text").textContent = r.text;
    $("ao-run").textContent = r ? "Reprendre" : "Démarrer";
    // Récupère l'espace des anciens passages (le navigateur ne recharge pas une
    // SPA : sans ça le stockage se remplirait). Garde le dernier passage + le journal.
    if (!running) { try { await natif.prune([cfg.Dernier, String(cfg.Dernier || "").replace(/\.txt$/, ".csv")], SOURCE); } catch (e) {} }
  }

  // ------------------------------------------------------------ branchements
  shadow.querySelectorAll("[data-group]").forEach(function (g) {
    g.addEventListener("click", function (e) {
      if (running) return; // réglages verrouillés pendant un passage
      var b = e.target.closest(".pill"); if (!b) return;
      cfg[g.dataset.group] = b.dataset.value; setPills(g.dataset.group, b.dataset.value); update();
    });
  });
  ["Prefix", "Keep", "Exclude"].forEach(function (k) { $("ao-" + k).addEventListener("input", function () { if (!running) update(); }); });
  ["Dedupe", "CopyAtEnd"].forEach(function (k) { $("ao-" + k).addEventListener("change", function () { if (!running) update(); }); });
  var collapsed = false;
  $("ao-min").onclick = function () { if (running) return; collapsed = !collapsed; $("ao-body").hidden = collapsed; $("ao-min").textContent = collapsed ? "+" : "—"; };
  // Pendant un passage, l'engrenage ramène à la vue compacte (le bouton Arrêter
  // doit rester accessible) au lieu d'ouvrir les paramètres.
  // Ouvrir les paramètres suspend le démarrage automatique (on ne lance pas un
  // passage sous les pieds de quelqu'un qui règle ou diagnostique) ; « Démarrer
  // ici » reste possible, et le démarrage auto revient en quittant la vue.
  function settingsOpen() { return !$("ao-set-view").hidden; }
  $("ao-gear").onclick = function () { if (running) { view("compact"); return; } var set = $("ao-set-view").hidden; if (set) cancelCountdown(); view(set ? "set" : "run"); };
  // Croix : ferme le panneau pour cette page seulement ; il revient au prochain
  // chargement ou rafraîchissement. Fermer coupe aussi le démarrage automatique
  // (sinon un passage partirait sans panneau pour l'arrêter) ; un passage en cours
  // s'arrête aussitôt (la recherche en cours est abandonnée, refaite à la reprise),
  // journal gardé : « Reprendre ici » au prochain chargement.
  $("ao-close").onclick = function () {
    closed = true;
    cancelCountdown();
    try { if (root.AOArmement) { root.AOArmement.suspend(); root.AOArmement.disarm(); } } catch (e) {}
    if (running && moteur) moteur.stop = true;
    try { if (mountObs) mountObs.disconnect(); } catch (e) {}
    host.remove();
  };
  $("ao-reset").onclick = function () { cfg = Moteur.defaults(""); cfg.Source = SOURCE; applyConfig(); update(); };

  // Un envoi Publipostage en cours dans N'IMPORTE QUELLE fenêtre ou onglet (marqueur
  // « pp-actif » rafraîchi pendant l'envoi, périmé après 30 s) : pas de relevé.
  async function envoiAilleurs() {
    try { var o = await api.storage.local.get("pp-actif"), m = o["pp-actif"]; return !!(m && Date.now() - m.t < 30000); } catch (e) { return false; }
  }
  function rendreVerrou() { if (root.__smVerrou === "releve") root.__smVerrou = null; }
  async function lancer(fn) {
    if (running) return;
    if ((root.PP && root.PP.moteurActif) || await envoiAilleurs()) { notice("Un envoi Publipostage est en cours : mets-le en pause avant de relever des adresses.", "warn"); rendreVerrou(); return; }
    readConfig();
    if (await runActiveElsewhere()) { notice("Un passage « " + Moteur.sourceName(cfg) + " » semble déjà en cours (autre onglet, ou page tout juste rechargée). Réessaie dans quelques secondes.", "warn"); rendreVerrou(); return; }
    if (closed) { rendreVerrou(); return; } // fermé pendant la vérification du verrou : on ne démarre pas
    heartbeat();
    $("ao-run").disabled = true; $("ao-export").disabled = true; $("ao-copy").disabled = true; $("ao-diag").disabled = true;
    try { await fn(); }
    catch (e) { ui.fin("Erreur : " + (e && e.message ? e.message : e), "warn"); }
  }
  // « Démarrer ici » : on arme SANS condition le champ où est le curseur (le clic
  // fait foi, pas de vérification « est-ce le bon champ »), puis on tape dedans.
  // « Démarrer » (ou « Reprendre ») depuis n'importe quelle page d'Outlook :
  //  - curseur déjà dans un champ du message (À, Cc…) : on prend ce champ ;
  //  - sinon le champ « À » du message ouvert ;
  //  - sinon on ouvre un nouveau message (même geste que le Publipostage).
  // Sur Gmail (pas de Publipostage) : le champ où est le curseur, comme avant.
  var demarrage = false;
  // Le relevé ne tape QUE dans un champ destinataire (À, Cc, Cci) VIDE : jamais
  // dans le corps, l'objet, la recherche, une discussion Teams, ni dans un message
  // que l'utilisateur a commencé.
  var LIB_DEST = /^(à|a|to|cc|cci|bcc|ccn|cco|destinataires|to recipients|an|para|aan)$/i;
  function estDestinataire(el) {
    if (!el || el.nodeType !== 1) return false;
    if (el.closest && el.closest('[data-testid="ao-panel"],[data-pp-panneau],[role="search"]')) return false;
    var O = root.PP && root.PP.Outlook;
    try { if (O && O.estChampDestinataire && O.estChampDestinataire(el)) return true; } catch (e) {}
    var lec = root.AOLecture && root.AOLecture.SEL && root.AOLecture.SEL[SOURCE];
    var sels = (lec && lec.toField) || [];
    for (var i = 0; i < sels.length; i++) { try { if (el.matches(sels[i])) return true; } catch (e) {} }
    return LIB_DEST.test((el.getAttribute("aria-label") || "").trim());
  }
  function estVide(el) {
    var O = root.PP && root.PP.Outlook;
    if (O && O.champVide) { try { return O.champVide(el); } catch (e) {} }
    var libre = el.value != null && (el.tagName === "INPUT" || el.tagName === "TEXTAREA") ? el.value : (el.textContent || "");
    if (libre.replace(/[\s\u200B-\u200D\u2060\uFEFF]/g, "")) return false;
    var ligne = el.closest && (el.closest('div[name="to"],[role="group"]') || el.parentElement);
    return !(ligne && ligne.querySelector('[data-hovercard-id],[contenteditable="false"],[role="listitem"]'));
  }
  function champActif() {
    var el = document.activeElement;
    return estDestinataire(el) ? el : null;
  }
  var REFUS_PLEIN = "Ce champ contient déjà des destinataires ou du texte : je n'y touche pas. Ouvrez un nouveau message vide, puis « Démarrer ».";
  async function preparerChamp() {
    var A = root.AOArmement, O = root.PP && root.PP.Outlook;
    if (!A) return false;
    A.disarm();
    var act = champActif();
    if (act) {
      if (!estVide(act)) { notice(REFUS_PLEIN, "warn"); return false; }
      return A.armHere();
    }
    if (!O || !O.champA) { notice("Cliquez dans le champ " + CHAMP + " (vide) d'un nouveau message, puis « Démarrer ».", "warn"); return false; }
    var champ = O.champA();
    if (champ) {
      // Une rédaction est déjà affichée : seulement si elle est vierge (sinon c'est
      // le message de l'utilisateur, on n'y touche pas).
      if (!O.redactionVide || !O.redactionVide()) {
        notice("Un message est en cours d'écriture (destinataire ou objet déjà saisi) : je n'y touche pas. Fermez-le — il reste dans vos brouillons — puis « Démarrer ».", "warn");
        return false;
      }
    } else {
      armHint("Ouverture d'un nouveau message…", "Je me place ensuite dans le champ « À ». Rien ne sera envoyé.");
      var r = await O.ouvrirRedaction();
      if (!r.ok) { notice(r.raison, "warn"); return false; }
      champ = O.champA();
      if (!champ) { notice("Champ « À » introuvable dans le nouveau message.", "warn"); return false; }
    }
    if (!estDestinataire(champ) || !estVide(champ)) { notice(REFUS_PLEIN, "warn"); return false; }
    try { champ.focus(); } catch (e) {}
    if (document.activeElement !== champ && !(champ.contains && champ.contains(document.activeElement))) {
      notice("Impossible de placer le curseur dans le champ « À ». Rien n'a été tapé.", "warn");
      return false;
    }
    return A.armHere();
  }
  $("ao-run").onclick = async function () {
    if (running || demarrage) return;
    if (root.__smVerrou && root.__smVerrou !== "releve") { notice("Un envoi Publipostage est en cours : mets-le en pause avant de relever des adresses.", "warn"); return; }
    root.__smVerrou = "releve";
    if ((root.PP && root.PP.moteurActif) || await envoiAilleurs()) { notice("Un envoi Publipostage est en cours : mets-le en pause avant de relever des adresses.", "warn"); return; }
    cancelCountdown(); view("run"); notice("", "");
    demarrage = true; $("ao-run").disabled = true;
    var ok = false;
    try { ok = await preparerChamp(); } catch (e) { notice("Démarrage : " + (e && e.message ? e.message : e), "warn"); }
    demarrage = false; $("ao-run").disabled = false;
    if (!ok && root.__smVerrou === "releve") root.__smVerrou = null;
    if (!ok) { if (!$("ao-notice").textContent) notice("Clique dans le champ " + CHAMP + " d'un message, puis « Démarrer ».", "warn"); hintPret(); return; }
    var r = resumable; lancer(function () { return moteur.passage(cfg, r); });
  };
  $("ao-fresh").onclick = async function () {
    await moteur.oublierReprise(); await refreshResume();
    notice("Le passage interrompu est mis de côté. Lancer repart de A.", "");
  };
  // Diagnostic : on vise le champ où est le curseur (« À »), on tape « a » pour
  // faire APPARAÎTRE la liste de suggestions, on la photographie, puis on efface.
  // Sans frappe, la liste de contacts n'est pas affichée : on ne verrait rien.
  $("ao-diag").onclick = async function () {
    if ((root.PP && root.PP.moteurActif) || running || await envoiAilleurs()) { notice("Un envoi ou un relevé est en cours : mettez-le en pause avant le diagnostic.", "warn"); return; }
    view("run"); notice("Diagnostic en cours… garde le message ouvert.", "");
    try {
      var armed = false;
      // Jamais de frappe hors d'un champ destinataire VIDE (corps, réponse, objet…).
      var cible = champActif();
      if (cible && estVide(cible)) { try { root.AOArmement.disarm(); armed = !!(root.AOArmement && root.AOArmement.armHere()); } catch (e) {} }
      await natif("diagnostic", { etape: armed ? "champ visé, avant frappe" : "aucun champ visé" });
      var lu = [];
      if (armed) {
        await natif("taper", { texte: "a" });
        await new Promise(function (r) { setTimeout(r, 1500); });
        await natif("diagnostic", { etape: "après avoir tapé a" });
        // Essai « Afficher plus de résultats » : on clique, puis on recapture pour
        // voir si la ligne a été trouvée et si la liste s'est agrandie.
        var essai = {};
        try { essai = await natif("plusEssai", {}); } catch (e) {}
        await new Promise(function (r) { setTimeout(r, 1500); });
        await natif("diagnostic", { etape: "après clic « plus de résultats » (ligne " + (essai && essai.trouve ? "TROUVÉE" : "non trouvée") + ", contacts avant " + (essai ? essai.avant : "?") + ")" });
        try { lu = await natif("lire", {}); } catch (e) {}
        await natif("effacer", { n: 3 });
        try { root.AOArmement.disarm(); } catch (e) {}
      }
      var path = await natif("diagnosticFichier", { lu: lu });
      notice((armed ? "" : "Clique d'abord dans le champ « À », puis relance. ") + "Diagnostic téléchargé : " + String(path).split("/").pop() + " (dans Téléchargements). Envoie-le-moi.", armed ? "ok" : "warn");
    } catch (e) { notice("Diagnostic : " + (e && e.message ? e.message : e), "warn"); }
  };
  // « Arrêter » coupe net : le moteur abandonne la recherche en cours (elle
  // sera refaite à la reprise), vide le champ « À » et rend la main en une
  // fraction de seconde. Le bouton se grise pour éviter les doubles clics.
  $("ao-stop").onclick = function () { moteur.stop = true; $("ao-stop").disabled = true; $("ao-live-mail").textContent = "Mise en pause…"; };
  // Crochet d'arrêt pour les tests bout en bout : la page (monde principal) émet
  // « ao-stop » sur document, que ce script de contenu reçoit (DOM partagé).
  document.addEventListener("ao-stop", function () { if (moteur) moteur.stop = true; });
  $("ao-export").onclick = async function () {
    var p = await moteur.exporter(cfg);
    if (p) notice(Moteur.plural(moteur.distinct.length, "adresse exportée", "adresses exportées") + " dans " + String(p).split("/").pop() + ".", "ok");
  };
  $("ao-copy").onclick = async function () {
    if (!moteur.distinct.length) return;
    await natif("copier", { texte: moteur.distinct.join("\n") });
    notice(Moteur.plural(moteur.distinct.length, "adresse copiée", "adresses copiées") + ".", "ok");
  };

  // ---- Import (historique ou reprise) / export du point de reprise.
  $("ao-import").onclick = function () { $("ao-file").click(); };
  $("ao-file").onchange = async function () {
    var inp = $("ao-file"), f = inp.files && inp.files[0]; if (!f) return;
    var text = ""; try { text = await f.text(); } catch (e) { notice("Lecture du fichier impossible.", "warn"); inp.value = ""; return; }
    inp.value = ""; // permet de réimporter le même fichier
    await importerTexte(text);
  };
  // Vider l'historique : la base accumulée ET le dernier passage (son fichier en
  // storage est enlevé par prune), puis recalcul du compteur / Export / Copier.
  $("ao-clear-base").onclick = async function () {
    if (running) { notice("Attends la fin du passage pour vider l'historique.", "warn"); return; }
    try {
      await api.storage.local.remove([BASE, DERNIER]);
      cfg.Dernier = "";
      if (moteur) moteur.distinct = [];
      try { await natif.prune([], SOURCE); } catch (e) {}
      await applyKnown();
      save();
      notice("Historique vidé : plus aucune adresse connue. Le prochain passage repart sur une base propre.", "ok");
    } catch (e) { notice("Impossible de vider l'historique : " + (e && e.message ? e.message : e), "warn"); }
  };
  $("ao-save-resume").onclick = async function () {
    try {
      var o = await api.storage.local.get(JOURNAL), j = o[JOURNAL];
      if (j && String(j).trim()) {
        var d = new Date(), p2 = function (x) { return String(x).padStart(2, "0"); };
        var nom = "reprise-" + SOURCE + "-" + d.getFullYear() + "-" + p2(d.getMonth() + 1) + "-" + p2(d.getDate()) + ".txt";
        await natif("enregistrerSous", { nom: nom, dossier: "", texte: String(j) });
        notice("Point de reprise enregistré : " + nom + " (dans Téléchargements). Réimporte-le pour continuer au même endroit.", "ok");
      } else { notice("Aucun passage en cours à reprendre. Utilise « Exporter » pour la liste d'adresses.", "warn"); }
    } catch (e) { notice("Export impossible : " + (e && e.message ? e.message : e), "warn"); }
  };

  // ---- Armement : l'utilisateur tape dans « À », une liste à adresses apparaît,
  //      et on démarre tout seul (compte à rebours annulable). C'est ce qui rend
  //      l'outil indépendant de la version d'Outlook.
  var countdownTimer = null;
  function armHint(title, msg) { var t = $("ao-arm-title"), m = $("ao-arm-msg"); if (t) t.textContent = title; if (m) m.textContent = msg; }
  function cancelCountdown() { if (countdownTimer) { clearTimeout(countdownTimer); countdownTimer = null; } }
  // Smart Mail : pas de démarrage automatique quand l'onglet « Envoyer » est
  // affiché ou qu'un envoi tourne (le Publipostage tape lui-même dans « À »).
  function megaBloque() { try { return !!(root.MEGA && !root.MEGA.relevePermis()); } catch (e) { return false; } }
  // Smart Mail : plus de démarrage automatique (compte à rebours) quand on
  // clique dans « À » : c'est le bouton « Démarrer » qui lance. Le repère
  // data-ao-armed reste posé pour les tests.
  function hintPret() {
    if (running) return;
    if (root.PP && root.PP.Outlook) armHint("Prêt", "Clique « Démarrer » depuis n'importe quelle page d'Outlook : j'ouvre un nouveau message et je relève les adresses que propose le champ « À ». Rien n'est envoyé.");
    else armHint("Prêt", "Clique dans le champ " + CHAMP + " d'un message, puis « Démarrer ». Rien n'est envoyé.");
  }
  function onArmed() {
    if (closed || running || megaBloque()) return;
    host.dataset.aoArmed = "1";
  }
  function onDisarmed() { cancelCountdown(); try { delete host.dataset.aoArmed; } catch (e) {} if (!running && !demarrage) hintPret(); }
  try { if (root.AOArmement) { root.AOArmement.onArmed(onArmed); root.AOArmement.onDisarmed(onDisarmed); } } catch (e) {}

  // ------------------------------------------------------------ démarrage
  (async function () {
    var saved = null;
    try { var o = await api.storage.local.get("reglages"); saved = o.reglages ? JSON.parse(o.reglages) : null; } catch (e) { saved = null; }
    cfg = Moteur.merge(Moteur.defaults(""), saved);
    cfg.Source = SOURCE; // la messagerie est celle de l'onglet, pas un choix
    // « Dernier fichier » (Exporter/Copier au prochain lancement) dans sa propre
    // clé par messagerie — jamais celui d'une autre messagerie, jamais écrasé.
    try { var dj = await api.storage.local.get(DERNIER); cfg.Dernier = dj[DERNIER] || ""; } catch (e) { cfg.Dernier = ""; }
    moteur = new Moteur(natif, ui, waitFn, JOURNAL);
    applyConfig(); update(); view("run"); hintPret();
    await refreshResume(); // nettoie aussi le stockage (prune)
    host.dataset.aoReady = "1"; // repère pour les tests : amorçage terminé
    // Smart Mail (mega.js) : adresses connues (historique + dernier passage,
    // comme « Copier ») et état du passage.
    root.AOPanneau = {
      host: host,
      adresses: async function () { return (await applyKnown()).slice(); },
      enCours: function () { return running; },
      notice: notice
    };
    try { document.dispatchEvent(new CustomEvent("mega-pret")); } catch (e) {}
  })();
})(typeof globalThis !== "undefined" ? globalThis : this);
