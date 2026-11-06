// Smart Mail : la partie « Adresses » n'a pas besoin du service worker (tout se
// passe dans le script de contenu) ; la partie « Publipostage » s'en sert pour
// écrire « Déjà envoyés » dans Téléchargements et pour l'envoi en arrière-plan.

// Bouton de l'extension : affiche ou masque le panneau dans l'onglet (mega.js).
chrome.action.onClicked.addListener(function (tab) {
  if (tab && tab.id != null) chrome.tabs.sendMessage(tab.id, { mega: "basculer" }).catch(function () {});
});

// Fichier texte écrit dans Téléchargements, en remplaçant le précédent du même nom
// (liste « Déjà envoyés » : publipostage-deja-envoyes.txt).
chrome.runtime.onMessage.addListener(function (m, envoyeur, repondre) {
  if (!m || m.pp !== "fichier" || typeof m.texte !== "string" || !/^[\w.-]+\.txt$/.test(m.nom || "")) return;
  var url = "data:text/plain;charset=utf-8," + encodeURIComponent(m.texte);
  chrome.downloads.download({ url: url, filename: m.nom, conflictAction: "overwrite", saveAs: false }, function (id) {
    var err = chrome.runtime.lastError;
    repondre(err || id == null ? { ok: false, raison: err ? err.message : "téléchargement refusé" } : { ok: true });
  });
  return true;
});

// Envoi en arrière-plan. Chrome ralentit les minuteries d'un onglet caché (un
// réveil par minute après 5 min) : l'onglet Outlook demande ici ses réveils, qui
// ne sont pas ralentis (attendre, dans outlook.js). Pendant l'envoi, l'onglet
// n'est pas déchargé par l'économiseur de mémoire et le Mac ne se met pas en
// veille (l'écran, lui, peut s'éteindre).
chrome.runtime.onMessage.addListener(function (m, envoyeur, repondre) {
  if (!m) return;
  if (m.pp === "reveil") {
    setTimeout(function () { repondre(true); }, Math.min(Math.max(0, Number(m.ms) || 0), 20000));
    return true;
  }
  if (m.pp === "envoi") {
    if (m.actif) chrome.power.requestKeepAwake("system"); else chrome.power.releaseKeepAwake();
    if (envoyeur.tab && envoyeur.tab.id != null) chrome.tabs.update(envoyeur.tab.id, { autoDiscardable: !m.actif }).catch(function () {});
    repondre(true);
  }
});
