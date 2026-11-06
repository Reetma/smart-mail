// Smart Mail : réunit les panneaux « Adresses » (relever les adresses que
// propose le champ À) et « Publipostage » (envoyer le message à chacune) en un
// seul panneau à deux onglets. Chaque panneau garde son code et son stockage ;
// ce fichier ne fait que :
//   - poser la barre d'onglets en haut de chacun, n'en montrer qu'un à la fois,
//     au même endroit (la position suit quand on déplace l'un ou l'autre) ;
//   - tenir l'annuaire commun : chaque personne relevée y entre en direct, et
//     l'onglet Envoyer s'en sert comme destinataires ;
//   - empêcher l'un de démarrer pendant que l'autre tourne : le relevé tape dans
//     « À » pour faire apparaître les suggestions, l'envoi y colle les adresses.
// Dernier script chargé, cadre principal d'Outlook uniquement (pas de Gmail :
// le Publipostage ne sait envoyer que depuis Outlook).

(function (root) {
  "use strict";
  if (root.top !== root.self) return;
  var api = root.chrome || root.browser;
  var CLE_ONGLET = "mega-onglet";

  var onglet = "adresses"; // onglet affiché : "adresses" ou "envoi"
  var masque = false;      // panneau masqué par le bouton de l'extension ou la croix
  var ao = null, pp = null; // { host, ombre } de chaque panneau, une fois prêts

  function envoiActif() { return !!(root.PP && root.PP.moteurActif); }
  function releveActif() { return !!(root.AOPanneau && root.AOPanneau.enCours()); }

  root.MEGA = {
    onglet: function () { return onglet; },
    // Démarrage automatique du relevé : seulement onglet « Relever » affiché et
    // aucun envoi en cours (lu par adresses/panel.js à chaque armement).
    relevePermis: function () { return !masque && onglet === "adresses" && !envoiActif(); }
  };

  var CSS = `
  .mega-tabs { display:flex;gap:4px;padding:8px 10px 0;background:var(--card);border-bottom:1px solid var(--line); }
  .mega-tab { flex:1;border:1px solid transparent;border-bottom:none;background:none;color:var(--muted);font:inherit;font-weight:600;font-size:12.5px;
    padding:7px 8px;border-radius:9px 9px 0 0;cursor:pointer;white-space:nowrap;position:relative;top:1px; }
  .mega-tab:hover { color:var(--ink2); }
  .mega-tab[aria-selected="true"] { background:var(--bg);color:var(--brand);border-color:var(--line); }
  .mega-tab .pastille { display:inline-block;width:7px;height:7px;border-radius:50%;background:#16A34A;margin-left:6px;vertical-align:1px; }
  .mega-tab .pastille[hidden] { display:none; }
  .mega-send { width:100%;margin:-2px 0 10px;display:inline-flex;align-items:center;justify-content:center;gap:8px;
    background:var(--brand-soft);color:var(--brand);border:1px solid transparent;border-radius:12px;padding:10px 16px;font-weight:700;
    transition:background .12s ease,border-color .12s ease; }
  .mega-send:hover:not(:disabled) { background:var(--brand-soft);border-color:var(--brand); }
  .mega-send:disabled { background:transparent;border:1px dashed var(--line);color:var(--faint);opacity:1; }
  .mega-send .compte { background:var(--brand);color:#fff;border-radius:999px;padding:1px 8px;font-size:11.5px;font-weight:700;min-width:20px;text-align:center; }
  .mega-send:disabled .compte { display:none; }
  .mega-send .fleche { display:inline-block;transition:transform .15s ease; }
  .mega-send:hover:not(:disabled) .fleche { transform:translateX(3px); }
  `;

  function barre(ombre, actif) {
    var st = document.createElement("style"); st.textContent = CSS; ombre.appendChild(st);
    var b = document.createElement("div"); b.className = "mega-tabs"; b.setAttribute("role", "tablist");
    [["adresses", "① Relever les adresses"], ["envoi", "② Envoyer"]].forEach(function (t) {
      var x = document.createElement("button");
      x.className = "mega-tab"; x.setAttribute("role", "tab"); x.dataset.onglet = t[0];
      x.setAttribute("aria-selected", String(t[0] === actif));
      x.appendChild(document.createTextNode(t[1]));
      var p = document.createElement("span"); p.className = "pastille"; p.hidden = true; p.title = "En cours";
      x.appendChild(p);
      x.addEventListener("click", function () { montrer(t[0]); });
      b.appendChild(x);
    });
    var panel = ombre.querySelector(".panel");
    panel.insertBefore(b, panel.firstChild);
    // Les boutons des panneaux ne volent pas le focus (champ « À » armé) : idem ici.
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
  }

  // Même position pour les deux panneaux : celui qu'on affiche prend la place
  // de celui qu'on quitte (déplacé à la main ou non).
  function copierPosition(de, vers) {
    ["top", "left", "right"].forEach(function (k) { vers.style[k] = de.style[k]; });
  }

  function montrer(o, sansSauver) {
    if (!ao || !pp) return;
    var avant = onglet === "envoi" ? pp.host : ao.host;
    onglet = o;
    var apres = o === "envoi" ? pp.host : ao.host;
    if (apres !== avant) copierPosition(avant, apres);
    ao.host.style.display = !masque && o === "adresses" ? "" : "none";
    pp.host.style.display = !masque && o === "envoi" ? "" : "none";
    [ao.ombre, pp.ombre].forEach(function (om) {
      om.querySelectorAll(".mega-tab").forEach(function (x) { x.setAttribute("aria-selected", String(x.dataset.onglet === o)); });
    });
    // Hors de l'onglet « Relever », le champ « À » ne déclenche plus de relevé.
    try {
      var A = root.AOArmement;
      if (A && !releveActif()) {
        if (root.MEGA.relevePermis()) A.resume(); else { A.suspend(); A.disarm(); }
      }
    } catch (e) {}
    if (!sansSauver) try { var s = {}; s[CLE_ONGLET] = o; api.storage.local.set(s); } catch (e) {}
  }

  function basculer() {
    masque = !masque;
    if (masque && releveActif()) { var stop = ao && ao.ombre.getElementById("ao-stop"); if (stop && !stop.disabled) stop.click(); }
    montrer(onglet, true);
  }

  // Pastille verte sur l'onglet de ce qui tourne, et armement coupé pendant un envoi.
  var etaitEnvoi = false;
  function veiller() {
    if (!ao || !pp) return;
    var r = releveActif(), e = envoiActif();
    [ao.ombre, pp.ombre].forEach(function (om) {
      om.querySelectorAll(".mega-tab").forEach(function (x) {
        x.querySelector(".pastille").hidden = !(x.dataset.onglet === "adresses" ? r : e);
      });
    });
    if (e !== etaitEnvoi) {
      etaitEnvoi = e;
      if (!e && masque) {
        var n = pp.ombre.getElementById("pp-notice");
        if (n && /warn/.test(n.className)) { masque = false; onglet = "envoi"; }
      }
      montrer(onglet, true);
    }
  }

  // ------------------------------------------------------------ annuaire commun
  // Toutes les personnes relevées, quelle que soit la messagerie : l'historique
  // d'« Adresses » (base-outlook, base-gmail) plus chaque adresse trouvée par le
  // relevé en cours, AU MOMENT où elle s'affiche (adresses/panel.js appelle
  // MEGA.trouve). L'onglet Envoyer s'en sert directement comme destinataires
  // (case « Annuaire commun »), sans copier-coller ni transfert.
  // « Vider l'historique » (⚙ d'Adresses) le vide aussi.
  var CLE_ANNUAIRE = "mega-annuaire", CLE_ANNUAIRE_ACTIF = "mega-annuaire-actif", BASES = ["base-outlook", "base-gmail"];
  var annuaire = new Set(), annuaireActif = true, annuaireCharge = false;
  var tPousser = null, tSauver = null;

  function lignes(v) { return v ? String(v).split(/\r?\n/) : []; }
  var ADRESSE_VALIDE = /^[a-z0-9._%+-]+@[a-z0-9-]+(?:\.[a-z0-9-]+)*\.[a-z]{2,}$/;
  function ajouter(mail) {
    var k = String(mail || "").trim().toLowerCase().replace(/^[.']+|[.']+$/g, "");
    if (!k || !ADRESSE_VALIDE.test(k) || /\.\.|^\.|\.@/.test(k) || annuaire.has(k)) return false;
    annuaire.add(k); return true;
  }
  function listeAnnuaire() { return Array.from(annuaire).sort(); }
  function annuaireChange() {
    clearTimeout(tPousser); tPousser = setTimeout(pousser, 200);
    clearTimeout(tSauver); tSauver = setTimeout(sauverAnnuaire, 1000);
  }
  function sauverAnnuaire() {
    if (!annuaireCharge) return;
    try { var o = {}; o[CLE_ANNUAIRE] = listeAnnuaire().join("\n"); api.storage.local.set(o); } catch (e) {}
  }
  // Affiche le compte des deux côtés et met à jour la liste de l'envoi.
  function pousser() {
    var n = annuaire.size;
    if (ao) {
      var b = ao.ombre.getElementById("mega-send");
      if (b) { b.querySelector(".compte").textContent = n.toLocaleString("fr-FR"); b.dataset.n = String(n); b.disabled = !n; b.title = n ? n.toLocaleString("fr-FR") + " personne(s) dans l'annuaire commun, utilisé par l'onglet ② Envoyer" : "L'annuaire se remplit pendant le relevé"; }
    }
    var P = root.PP && root.PP.Panneau;
    if (P && annuaireCharge) P.majAnnuaire(listeAnnuaire(), annuaireActif);
  }
  root.MEGA.trouve = function (mail) { if (ajouter(mail)) annuaireChange(); };
  root.MEGA.annuaire = function () { return listeAnnuaire(); };

  async function chargerAnnuaire() {
    var o = {};
    try { o = await api.storage.local.get([CLE_ANNUAIRE, CLE_ANNUAIRE_ACTIF].concat(BASES)); } catch (e) {}
    lignes(o[CLE_ANNUAIRE]).forEach(ajouter);
    BASES.forEach(function (k) { lignes(o[k]).forEach(ajouter); });
    if (typeof o[CLE_ANNUAIRE_ACTIF] === "boolean") annuaireActif = o[CLE_ANNUAIRE_ACTIF];
    annuaireCharge = true;
    annuaireChange();
  }
  // Historique d'Adresses modifié (fin de relevé, import, ou relevé fait dans
  // un onglet Gmail) : on complète. Historique vidé : on repart des historiques
  // qui restent (les trouvailles en direct non enregistrées partent avec).
  try {
    api.storage.onChanged.addListener(function (ch, zone) {
      if (zone !== "local") return;
      var vide = false, plus = false;
      BASES.forEach(function (k) {
        if (!ch[k]) return;
        if (ch[k].newValue == null) vide = true;
        else lignes(ch[k].newValue).forEach(function (m) { if (ajouter(m)) plus = true; });
      });
      if (vide) {
        api.storage.local.get(BASES).then(function (o) {
          annuaire.clear();
          BASES.forEach(function (k) { lignes(o[k]).forEach(ajouter); });
          annuaireChange();
        });
      } else if (plus) annuaireChange();
    });
  } catch (e) {}

  // Côté Relever : un raccourci vers l'onglet Envoyer, l'annuaire y est déjà.
  function boutonEnvoyer() {
    var demarrer = ao.ombre.getElementById("ao-run");
    if (!demarrer) return;
    var b = document.createElement("button");
    b.className = "btn mega-send"; b.id = "mega-send"; b.disabled = true;
    [["libelle", "Envoyer à l'annuaire"], ["compte", "0"], ["fleche", "→"]].forEach(function (p) {
      var x = document.createElement("span"); x.className = p[0]; x.textContent = p[1]; b.appendChild(x);
    });
    b.title = "Les personnes relevées sont déjà dans l'annuaire commun, utilisé par l'onglet ② Envoyer";
    demarrer.parentNode.after(b); // sous « Démarrer »
    b.addEventListener("click", function () {
      if (!annuaireActif) { annuaireActif = true; choixAnnuaire(true); pousser(); }
      montrer("envoi");
    });
  }
  function choixAnnuaire(actif) {
    annuaireActif = actif;
    try { var o = {}; o[CLE_ANNUAIRE_ACTIF] = actif; api.storage.local.set(o); } catch (e) {}
  }

  function brancher(host) {
    // La croix d'« Adresses » retirait son panneau pour de bon : ici elle masque
    // le panneau Smart Mail (le bouton de l'extension le fait revenir).
    var fermer = host.shadowRoot.getElementById("ao-close");
    if (fermer) {
      fermer.title = "Masquer (le bouton de l'extension le fait revenir)";
      fermer.addEventListener("click", function (e) { e.stopImmediatePropagation(); e.preventDefault(); basculer(); }, true);
    }
  }

  // Les deux panneaux finissent leur amorçage (lecture du stockage) chacun à
  // son rythme : on attend les deux.
  var fini = false;
  async function essayer() {
    if (fini) return;
    var hA = document.querySelector('[data-testid="ao-panel"]'), hP = document.querySelector("[data-pp-panneau]");
    if (!hA || !hP || !root.AOPanneau || !(root.PP && root.PP.Panneau) || !hA.shadowRoot || !hP.shadowRoot) return;
    fini = true;
    ao = { host: hA, ombre: hA.shadowRoot };
    pp = { host: hP, ombre: hP.shadowRoot };
    hP.style.right = "12px"; // le Publipostage s'était décalé à gauche d'Adresses
    var o = null;
    try { o = (await api.storage.local.get(CLE_ONGLET))[CLE_ONGLET]; } catch (e) {}
    // Un envoi en cours de route s'affiche d'abord (on y revient pour le reprendre).
    if (o !== "adresses" && o !== "envoi") o = "adresses";
    if (root.PP.Panneau && pp.ombre.getElementById("pp-reprise") && !pp.ombre.getElementById("pp-reprise").hidden) o = "envoi";
    onglet = o;
    barre(ao.ombre, o); barre(pp.ombre, o);
    brancher(hA);
    boutonEnvoyer();
    root.PP.Panneau.surChoixAnnuaire = choixAnnuaire;
    montrer(o, true);
    await chargerAnnuaire();
    setInterval(veiller, 700);
    document.documentElement.dataset.megaPret = "1"; // repère pour les tests
  }
  document.addEventListener("mega-pret", essayer);
  var tentatives = 0, t = setInterval(function () { essayer(); if (fini || ++tentatives > 60) clearInterval(t); }, 500);
  essayer();

  try {
    api.runtime.onMessage.addListener(function (m) { if (m && m.mega === "basculer") basculer(); });
  } catch (e) {}
})(typeof globalThis !== "undefined" ? globalThis : this);
