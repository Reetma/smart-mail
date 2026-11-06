// Émulation d'un onglet Chrome en arrière-plan, à injecter dans UN monde JS de
// la page (monde de la page via addInitScript, monde isolé de l'extension via
// CDP Runtime.evaluate). Chromium headless rend toutes les pages « visibles » et
// Playwright désactive le bridage des onglets cachés : on reproduit donc ici ce
// que fait Chrome pour un onglet caché, monde par monde :
//  - API de visibilité : document.visibilityState / hidden / hasFocus() ;
//  - minuteries : un réveil par seconde au plus (mode « seconde ») ; après 5 min
//    cachées, les minuteries en chaîne ne sont plus réveillées qu'une fois par
//    minute (mode « minute », « intensive wake up throttling ») : on y range
//    setInterval (chaque répétition compte dans la chaîne) ; les setTimeout restent
//    à 1/s (hypothèse prudente : en vrai ceux créés depuis une minuterie passent
//    aussi à 1/min, ce qui ne ferait que ralentir l'extension) ;
//  - requestAnimationFrame : suspendu tant que l'onglet est caché (monde de la page) ;
//  - horloge : décalage de Date (veille du Mac : l'heure murale saute, les
//    minuteries ne comptent pas le temps de veille) ;
//  - Math.random à graine fixe (monde de l'extension : tests déterministes).
// MessageChannel et les réponses du service worker ne sont pas bridés (comme Chrome).
function __emuler(graine, monde) {
  var g = globalThis;
  if (g.__AP) return "deja";
  var RDate = Date, rnow = RDate.now.bind(RDate);
  var RST = g.setTimeout.bind(g), RCT = g.clearTimeout.bind(g);
  var RRAF = g.requestAnimationFrame ? g.requestAnimationFrame.bind(g) : null;
  var RCAF = g.cancelAnimationFrame ? g.cancelAnimationFrame.bind(g) : null;
  var ctl = { monde: monde, cache: false, focus: true, mode: "aucun", origine: rnow(), decalage: 0, tirs: 0 };
  g.__AP = ctl;

  // ---------------------------------------------------------------- horloge
  function FD() {
    var a = [].slice.call(arguments);
    if (!new.target) return new RDate(rnow() + ctl.decalage).toString();
    if (!a.length) return new RDate(rnow() + ctl.decalage);
    return new (Function.prototype.bind.apply(RDate, [null].concat(a)))();
  }
  FD.prototype = RDate.prototype;
  FD.now = function () { return rnow() + ctl.decalage; };
  FD.parse = RDate.parse; FD.UTC = RDate.UTC;
  g.Date = FD;

  // ---------------------------------------------------------------- hasard
  if (graine != null) {
    var s = graine >>> 0;
    Math.random = function () {
      s = (s + 0x6D2B79F5) >>> 0; var t = s;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // ---------------------------------------------------------------- visibilité
  try {
    var DP = Document.prototype;
    var dVis = Object.getOwnPropertyDescriptor(DP, "visibilityState");
    var dHid = Object.getOwnPropertyDescriptor(DP, "hidden");
    var hf = DP.hasFocus;
    Object.defineProperty(DP, "visibilityState", { configurable: true, enumerable: true, get: function () { return ctl.cache ? "hidden" : dVis.get.call(this); } });
    Object.defineProperty(DP, "hidden", { configurable: true, enumerable: true, get: function () { return ctl.cache ? true : dHid.get.call(this); } });
    DP.hasFocus = function () { return ctl.cache || !ctl.focus ? false : hf.call(this); };
  } catch (e) {}

  // ---------------------------------------------------------------- minuteries
  var seq = 0, actifs = new Map();
  function cible(t, du) {
    if (ctl.mode === "aucun") return du;
    if (ctl.mode === "minute" && t.repete) return ctl.origine + Math.ceil((du - ctl.origine) / 60000) * 60000;
    return Math.ceil(du / 1000) * 1000;
  }
  function armer(t, du) {
    t.du = du; t.cible = cible(t, du);
    t.h = RST(function () { tirer(t); }, Math.max(0, t.cible - rnow()));
  }
  function tirer(t) {
    if (!actifs.has(t.id)) return;
    var c = cible(t, t.du), n = rnow();
    if (c > n + 2) { t.cible = c; t.h = RST(function () { tirer(t); }, c - n); return; } // bridage durci entre-temps
    if (t.repete) { t.n++; armer(t, rnow() + t.ms); } else actifs.delete(t.id);
    ctl.tirs++;
    if (typeof t.cb !== "function") return;
    try { t.cb.apply(g, t.args); } catch (e) { RST(function () { throw e; }, 0); }
  }
  function relacher() {
    actifs.forEach(function (t) {
      var c = cible(t, t.du);
      if (Math.abs(c - t.cible) > 1) { RCT(t.h); t.cible = c; t.h = RST(function () { tirer(t); }, Math.max(0, c - rnow())); }
    });
  }
  g.setTimeout = function (cb, ms) {
    var t = { id: ++seq, cb: cb, args: [].slice.call(arguments, 2), ms: Math.max(0, Number(ms) || 0), repete: false, n: 0 };
    actifs.set(t.id, t); armer(t, rnow() + t.ms); return t.id;
  };
  g.setInterval = function (cb, ms) {
    var t = { id: ++seq, cb: cb, args: [].slice.call(arguments, 2), ms: Math.max(1, Number(ms) || 0), repete: true, n: 0 };
    actifs.set(t.id, t); armer(t, rnow() + t.ms); return t.id;
  };
  g.clearTimeout = g.clearInterval = function (id) { var t = actifs.get(id); if (t) { RCT(t.h); actifs.delete(id); } };

  // ---------------------------------------------------------------- rAF (monde de la page)
  if (RRAF && monde === "page") {
    var nr = 0, file = new Map();
    g.requestAnimationFrame = function (cb) { if (!ctl.cache) return RRAF(cb); var id = -(++nr); file.set(id, cb); return id; };
    g.cancelAnimationFrame = function (id) { if (id < 0) file.delete(id); else RCAF(id); };
    ctl.viderRaf = function () { var l = []; file.forEach(function (cb) { l.push(cb); }); file.clear(); l.forEach(function (cb) { RRAF(cb); }); };
    ctl.rafEnAttente = function () { return file.size; };
  }

  // ---------------------------------------------------------------- pilotage
  // o : { cache, focus, mode: "aucun"|"seconde"|"minute", decalage }. evenements :
  // émet visibilitychange / blur / focus (une seule fois, depuis le monde de la page).
  ctl.etat = function (o, evenements) {
    var avCache = ctl.cache, avFocus = ctl.focus;
    if (o && "mode" in o) { if (o.mode === "minute" && ctl.mode !== "minute") ctl.origine = rnow(); ctl.mode = o.mode; }
    if (o && "cache" in o) ctl.cache = !!o.cache;
    if (o && "focus" in o) ctl.focus = !!o.focus;
    if (o && "decalage" in o) ctl.decalage = Number(o.decalage) || 0;
    relacher();
    if (!ctl.cache && ctl.viderRaf) ctl.viderRaf();
    if (evenements) {
      if (avCache !== ctl.cache) document.dispatchEvent(new Event("visibilitychange"));
      var f = !ctl.cache && ctl.focus, fa = !avCache && avFocus;
      if (f !== fa) { try { window.dispatchEvent(new FocusEvent(f ? "focus" : "blur")); } catch (e) {} }
    }
    return { cache: ctl.cache, focus: ctl.focus, mode: ctl.mode, decalage: ctl.decalage };
  };
  return "ok";
}
