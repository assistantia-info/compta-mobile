/* Compta Mobile — configuration (V1)
 * Defaults prudents : dryRun = true (AUCUNE écriture OneDrive tant que non
 * désactivé dans Réglages). Réglages persistés en localStorage (cm-config).
 * Invariants verrouillés ici (décisions 2026-09-04) : lignes 7–88, colonnes
 * A–E, jamais F (case de Fabrice), jamais d'insertion structurelle (O1),
 * feuille 'essai'/Feuil5/dep-fixe/ECO = zero touch (O6).
 */
"use strict";
window.CM = window.CM || {};

/* Version affichee en Diagnostic (aligner sur le CACHE de sw.js a chaque deploy). */
CM.APP_VERSION = "21 (compta-shell-v21)";

CM.DEFAULT_CONFIG = {
  authMode: "link",   /* "link" (lien de partage, voie b) | "oauth" (Microsoft Graph, voie a) */
  dryRun: true,       /* SECURITE : true = aucune écriture réelle */
  shareUrl: "",       /* lien OneDrive « Peut modifier » (voie b) */
  oauthClientId: "8b7cced7-e3de-43e3-afda-07252457fc15",  /* app hub "Electra Hub Mail" (client public, non secret), pose par Fo le 14/11 2026 */
  /* 14/11/2026 (apres-midi) : cible = le VRAI classeur du compte outlook.com, partage
   * avec fabrice.atmo@live.fr par lien « Peut modifier » (testé : roles write). Mode
   * "share" : le lien est résolu en item (drives/{driveId}/items/{id}) via /shares.
   * Mode "path" (repli) : la copie API sur le drive SPO du compte (workbookPath). */
  workbookMode: "share",
  workbookShareUrl: "https://1drv.ms/x/c/78cebc42e25b035f/IQCo_ZtF7Pr6Q4yNKfETw1-EAaB3GKL9ls3JyFZyid0lLx0?e=MXzBSp",
  workbookPath: "Documents/Compta familiale API/depenses fixe sept 2026-2027 API.xlsx",
  journalSheet: "RELSEPT26",
  journalFirstRow: 7,  /* première ligne de saisie (dans les plages SUM E89/D90) */
  journalLastRow: 88,  /* dernière ligne AVANT les totaux — jamais en dessous */
  oauthScopes: ["Files.ReadWrite", "User.Read", "offline_access"]
};

CM.getConfig = function () {
  try {
    const raw = localStorage.getItem("cm-config");
    return raw ? Object.assign({}, CM.DEFAULT_CONFIG, JSON.parse(raw))
               : Object.assign({}, CM.DEFAULT_CONFIG);
  } catch (e) { return Object.assign({}, CM.DEFAULT_CONFIG); }
};

CM.saveConfig = function (cfg) {
  localStorage.setItem("cm-config", JSON.stringify(cfg));
  return cfg;
};