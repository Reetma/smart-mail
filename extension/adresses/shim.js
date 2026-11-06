// Premier fichier du script de contenu. Donne un nom commun à l'API de
// l'extension : Firefox et Safari exposent `browser` (promesses), Chrome et
// Edge exposent `chrome` (callbacks, mais `chrome.storage.*` rend déjà des
// promesses en MV3). `var` (et pas `const`) pour que les fichiers suivants du
// même script de contenu voient bien `api` dans l'espace global isolé.
var api = globalThis.browser || globalThis.chrome;
