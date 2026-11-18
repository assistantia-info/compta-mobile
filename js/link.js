/* Compta Mobile — adaptateur lien de partage OneDrive (voie b)
 * Un lien « Peut modifier » donne un accès pré-authentifié :
 *   GET  https://api.onedrive.com/v1.0/shares/{shareId}/driveItem  (métadonnées)
 *   GET  .../root/content                                          (octets du classeur)
 *   PUT  .../root/content + If-Match etag                          (écriture)
 * Lecture = officiellement documentée ; écriture PUT = fonctionnelle en
 * pratique sur OneDrive perso mais NON garantie par Microsoft (consigne du
 * readme projet) — d'où le test PUT dédié dans l'onglet Diagnostic.
 * Toute écriture passe par la chirurgie zéro perte (CM.xlsx).
 * shareId = "u!" + base64url du lien (encodage documenté Microsoft).
 */
"use strict";
window.CM = window.CM || {};
CM.link = (function () {
  const API = "https://api.onedrive.com/v1.0/shares/";

  function shareIdFromUrl(url) {
    const clean = String(url).trim().split("?")[0];
    const b64 = btoa(clean).replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_");
    return "u!" + b64;
  }
  async function getMeta(shareId) {
    const resp = await fetch(API + shareId + "/driveItem");
    if (!resp.ok) throw new Error("GET driveItem " + resp.status + " " + resp.statusText);
    return await resp.json();
  }
  async function getContent(shareId) {
    const resp = await fetch(API + shareId + "/root/content");
    if (!resp.ok) throw new Error("GET content " + resp.status + " " + resp.statusText);
    return new Uint8Array(await resp.arrayBuffer());
  }
  async function putContent(shareId, bytes, etag) {
    const resp = await fetch(API + shareId + "/root/content", {
      method: "PUT",
      headers: Object.assign(
        { "Content-Type": "application/octet-stream" },
        etag ? { "If-Match": etag } : {}
      ),
      body: bytes
    });
    if (!(resp.status === 200 || resp.status === 201)) {
      throw new Error("PUT content " + resp.status + " " + resp.statusText);
    }
    return await resp.json();
  }

  return {
    shareIdFromUrl: shareIdFromUrl,

    /* Diagnostic complet (critère b de reprise).
       putTest : re-PUT des octets identiques avec If-Match (validation écriture). */
    diagnose: async function (shareUrl, opts) {
      const out = { steps: [] };
      const log = function (k, ok, msg) { out.steps.push({ step: k, ok: ok, detail: msg }); };
      const sid = shareIdFromUrl(shareUrl);
      log("encodage", true, "shareId = " + sid);
      const meta = await getMeta(sid);
      out.meta = { name: meta.name, size: meta.size, etag: meta.eTag, modified: meta.lastModifiedDateTime };
      log("GET métadonnées", true, meta.name + " | " + meta.size + " octets | modifié " + meta.lastModifiedDateTime);
      const bytes = await getContent(sid);
      log("GET contenu", true, bytes.length + " octets reçus");
      const sig = String.fromCharCode(bytes[0], bytes[1]);
      if (sig !== "PK") throw new Error("contenu non zip (signature " + sig + ")");
      log("signature zip", true, "PK ok (" + CM.xlsx.parseCentral(bytes).length + " parts)");
      const cfg = CM.getConfig();
      const state = await CM.xlsx.journalState(bytes, cfg.journalSheet, cfg.journalFirstRow, cfg.journalLastRow);
      out.journal = state;
      log("journal", true, cfg.journalSheet + " : " + state.usedRows + " ligne(s) utilisée(s), première libre = " + state.firstFreeRow);
      if (opts && opts.putTest) {
        const res = await putContent(sid, bytes, meta.eTag);
        out.putOk = true;
        log("PUT test", true, "écriture acceptée, nouvel etag " + (res.eTag || "?"));
      }
      return out;
    },

    pushExpense: async function (shareUrl, exp, sheet, first, last) {
      const sid = shareIdFromUrl(shareUrl);
      const meta = await getMeta(sid);
      const bytes = await getContent(sid);
      const r = await CM.xlsx.insertExpense(bytes, sheet, exp, first, last);
      await putContent(sid, r.bytes, meta.eTag);
      return { row: r.row };
    },

    readJournalState: async function (shareUrl, sheet, first, last) {
      const sid = shareIdFromUrl(shareUrl);
      const bytes = await getContent(sid);
      return await CM.xlsx.journalState(bytes, sheet, first, last);
    }
  };
})();