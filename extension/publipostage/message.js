// Message : HTML collé nettoyé par liste blanche (DOMPurify : pas de script, pas
// d'attribut d'événement, liens http/https/mailto seulement), découpé en
// segments pour la frappe (texte avec retours à la ligne, liens), texte brut.
//
// Exposé dans window.PP.Message.

(function (root) {
  "use strict";
  var PP = (root.PP = root.PP || {});
  var DP = root.DOMPurify;

  var CONFIG = {
    ALLOWED_TAGS: ["a", "b", "strong", "i", "em", "u", "s", "strike", "br", "p", "div", "span", "font", "ul", "ol", "li", "blockquote",
      "h1", "h2", "h3", "h4", "hr", "table", "thead", "tbody", "tr", "td", "th"],
    ALLOWED_ATTR: ["href", "style", "color", "align", "colspan", "rowspan"],
    ALLOWED_URI_REGEXP: /^(?:https?:|mailto:)/i,
  };

  function nettoyer(html) {
    if (!DP || typeof DP.sanitize !== "function") throw new Error("Nettoyage HTML indisponible (DOMPurify absent).");
    return sansLiensVides(sansParagraphes(DP.sanitize(String(html || ""), CONFIG)));
  }
  // Lien sans texte (« <a href="ancienne-adresse"></a> », reste d'un lien remplacé) :
  // invisible dans l'éditeur, retiré par Outlook au collage, il arrêtait l'envoi
  // (« Lien absent du mail »). Il est défait : son contenu éventuel (<br>) reste.
  function sansLiensVides(html) {
    if (html.indexOf("<a") < 0 && html.indexOf("<A") < 0) return html;
    var doc = new DOMParser().parseFromString("<div id='pp-racine'>" + html + "</div>", "text/html");
    var racine = doc.getElementById("pp-racine");
    var vides = [].filter.call(racine.getElementsByTagName("a"), function (a) { return !a.textContent.replace(/[\s​-‍⁠]/g, ""); });
    if (!vides.length) return html;
    vides.forEach(function (a) {
      while (a.firstChild) a.parentNode.insertBefore(a.firstChild, a);
      a.remove();
    });
    return racine.innerHTML;
  }
  // Outlook ajoute une marge verticale aux <p> (une ligne vide entre les blocs),
  // que notre éditeur masque (.editeur p { margin:0 }) : le même message paraît
  // en simple interligne ici et en double dans Outlook. Les paragraphes deviennent
  // donc de simples <div> (sans marge, comme texteEnHtml), toute marge en style
  // retirée. Les lignes vides voulues (<p><br></p>, <p>&nbsp;</p>) restent des
  // <div> vides et comptent toujours pour une ligne vide.
  function sansParagraphes(html) {
    if (html.indexOf("<p") < 0 && html.indexOf("<P") < 0) return html;
    var doc = new DOMParser().parseFromString("<div id='pp-racine'>" + html + "</div>", "text/html");
    var racine = doc.getElementById("pp-racine");
    for (var ps = racine.getElementsByTagName("p"); ps.length;) {
      var p = ps[0], d = doc.createElement("div");
      for (var i = 0; i < p.attributes.length; i++) {
        var nom = p.attributes[i].name, val = p.attributes[i].value;
        if (nom === "style") { val = val.replace(/(?:margin|padding)[^;]*;?/gi, "").trim(); if (!val) continue; }
        d.setAttribute(nom, val);
      }
      while (p.firstChild) d.appendChild(p.firstChild);
      p.parentNode.replaceChild(d, p); // <p> ne peut pas contenir de <p> : pas d'imbrication
    }
    return racine.innerHTML;
  }
  function echapper(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  // Texte simple (une ligne = un paragraphe) en HTML.
  function texteEnHtml(t) {
    return String(t || "").replace(/\r\n?/g, "\n").split("\n")
      .map(function (l) { return l ? "<div>" + echapper(l) + "</div>" : "<div><br></div>"; }).join("");
  }

  var BLOCS = /^(P|DIV|UL|OL|H1|H2|H3|H4|BLOCKQUOTE|TABLE|THEAD|TBODY|TR)$/;

  // Segments à taper : { t: "texte", s } (les retours à la ligne sont des "\n")
  // et { t: "lien", href, texte, html }. Le gras, les listes, etc. deviennent du texte.
  function segments(html) {
    var doc = new DOMParser().parseFromString("<div>" + nettoyer(html) + "</div>", "text/html");
    var ops = [], cur = "";
    // Le dernier caractère émis est-il un retour à la ligne (ou est-on au tout début) ?
    function finLigne() {
      if (cur) return cur.slice(-1) === "\n";
      return !ops.length; // juste après un lien : non
    }
    function debut() { return cur === "" && !ops.length; }
    function texte(s) {
      if (!s) return;
      if (debut() || finLigne()) s = s.replace(/^ +/, "");
      cur += s;
    }
    function ligne() { if (!debut()) cur += "\n"; }
    function finDeBloc() { if (!debut() && !finLigne()) ligne(); }
    function vider() { if (cur) { ops.push({ t: "texte", s: cur }); cur = ""; } }
    function parcourir(n) {
      if (n.nodeType === 3) { texte(n.nodeValue.replace(/[\t\n\r ]+/g, " ").replace(/\u00a0/g, " ")); return; }
      if (n.nodeType !== 1) return;
      var tag = n.tagName;
      if (tag === "BR") { cur += "\n"; return; }
      if (tag === "HR") { finDeBloc(); ligne(); return; }
      if (tag === "A" && n.getAttribute("href")) {
        var href = n.getAttribute("href");
        var lib = (n.textContent || "").replace(/\s+/g, " ").trim() || href;
        vider();
        ops.push({ t: "lien", href: href, texte: lib, html: '<a href="' + echapper(href) + '">' + echapper(lib) + "</a>" });
        return;
      }
      // Paragraphe vide (« <p>&nbsp;</p> » des mails Outlook) : une ligne vide.
      if ((tag === "P" || tag === "DIV") && !n.querySelector("br,a,img") && !(n.textContent || "").replace(/[\s\u00a0]/g, "")) {
        finDeBloc(); ligne(); return;
      }
      if (tag === "LI") { finDeBloc(); texte("- "); }
      else if (BLOCS.test(tag)) finDeBloc();
      for (var c = n.firstChild; c; c = c.nextSibling) parcourir(c);
      if (tag === "TD" || tag === "TH") texte(" ");
      if (tag === "LI" || BLOCS.test(tag)) finDeBloc();
    }
    parcourir(doc.body.firstChild);
    vider();
    // Pas d'espaces ni de lignes vides à la fin, pas d'espaces en fin de ligne.
    for (var i = 0; i < ops.length; i++) if (ops[i].t === "texte") ops[i].s = ops[i].s.replace(/ +\n/g, "\n");
    while (ops.length && ops[ops.length - 1].t === "texte") {
      var d = ops[ops.length - 1];
      d.s = d.s.replace(/[\s]+$/, "");
      if (d.s) break;
      ops.pop();
    }
    return ops;
  }

  // Adresses web écrites en clair (https://…, http://…, www.…) : transformées en
  // liens cliquables, sauf si elles sont déjà dans un lien. La ponctuation qui
  // suit (point, virgule, parenthèse…) reste en dehors du lien.
  var URL_NUE = /\b(?:https?:\/\/|www\.)[^\s<>"']*[^\s<>"'.,;:!?)\]}»]/gi;
  function lienifier(html) {
    var doc = new DOMParser().parseFromString("<div>" + nettoyer(html) + "</div>", "text/html");
    var racine = doc.body.firstChild;
    var marche = doc.createTreeWalker(racine, 4); // NodeFilter.SHOW_TEXT
    var textes = [];
    for (var n = marche.nextNode(); n; n = marche.nextNode()) {
      if (!(n.parentElement && n.parentElement.closest("a"))) textes.push(n);
    }
    textes.forEach(function (t) {
      var s = t.nodeValue, m, dernier = 0;
      URL_NUE.lastIndex = 0;
      if (!URL_NUE.test(s)) return;
      URL_NUE.lastIndex = 0;
      var frag = doc.createDocumentFragment();
      while ((m = URL_NUE.exec(s))) {
        if (m.index > dernier) frag.appendChild(doc.createTextNode(s.slice(dernier, m.index)));
        var a = doc.createElement("a");
        a.setAttribute("href", /^www\./i.test(m[0]) ? "https://" + m[0] : m[0]);
        a.textContent = m[0];
        frag.appendChild(a);
        dernier = m.index + m[0].length;
      }
      if (dernier < s.length) frag.appendChild(doc.createTextNode(s.slice(dernier)));
      t.parentNode.replaceChild(frag, t);
    });
    return nettoyer(racine.innerHTML);
  }

  function texteBrut(html) {
    return segments(html).map(function (o) { return o.t === "lien" ? o.texte : o.s; }).join("");
  }
  function liens(html) {
    return segments(html).filter(function (o) { return o.t === "lien"; });
  }

  PP.Message = { nettoyer: nettoyer, lienifier: lienifier, texteEnHtml: texteEnHtml, segments: segments, texteBrut: texteBrut, liens: liens, echapper: echapper };
})(typeof globalThis !== "undefined" ? globalThis : this);
