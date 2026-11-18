/* Compta Mobile — adaptateur OAuth Microsoft Graph (voie a)
 * Flux SPA + PKCE (sans secret client), endpoint common (comptes perso + org).
 * Cible : le vrai classeur « depenses fixe sept 2026-2027.xlsx » partagé depuis le
 * compte outlook.com par lien « Peut modifier » (résolu en item via /shares), ou en
 * repli la copie API (workbookPath) du drive SPO du compte.
 * Écriture : PATCH range A{r}:E{r} sur la feuille journal (jamais F, jamais
 * d'insertion structurelle), puis calculate Full pour propager la chaîne
 * mensuelle (E7 = mois-1!E9 ...). Session de classeur persistChanges.
 */
"use strict";
window.CM = window.CM || {};
CM.graph = (function () {
  const AUTH = "https://login.microsoftonline.com/common/oauth2/v2.0/";
  const GRAPH = "https://graph.microsoft.com/v1.0/";
  const BETA = "https://graph.microsoft.com/beta/";
  const STORE_KEY = "cm-oauth-tokens";
  let _session = null, _sessionItemId = null, _sessionDriveId = null;

  /* Adressage classeur : me/drive/items/{id} par défaut,
   * drives/{driveId}/items/{id} pour un item partagé d'un autre compte. */
  function wbBase(itemId, driveId) {
    return driveId ? "drives/" + driveId + "/items/" + itemId : "me/drive/items/" + itemId;
  }

  function b64url(bytes) {
    let s = btoa(String.fromCharCode.apply(null, bytes));
    return s.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }
  async function pkce() {
    const v = new Uint8Array(64);
    crypto.getRandomValues(v);
    const verifier = b64url(v);
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
    return { verifier: verifier, challenge: b64url(new Uint8Array(digest)) };
  }
  /* Session persistee (localStorage) : l'app rouvre deja connectee ;
   * renouvellement silencieux par refresh_token (offline_access). */
  function loadTokens() {
    try { return JSON.parse(localStorage.getItem(STORE_KEY) || "null"); } catch (e) { return null; }
  }
  function saveTokens(t) { localStorage.setItem(STORE_KEY, JSON.stringify(t)); }

  async function tokenRequest(body) {
    const resp = await fetch(AUTH + "token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(body).toString()
    });
    const j = await resp.json();
    if (!resp.ok) throw new Error("token : " + (j.error_description || j.error));
    return j;
  }

  /* ---- Ventilation (chantier Fabrice 14/11) : report de l'écriture du journal
   * dans les tableaux du mois via les plages nommées posées le 14/11
   * (VentCarburantSep26... = <mois>!F33:J43, VentPharmacieSep26... = <mois>!F50:H55).
   * L'échec de ventilation n'est JAMAIS un échec d'entrée : l'appelant journalise. */
  const VENT_MONTHS = {
    SEPTEMBRE26: "Sep26", OCTOBRE26: "Oct26", NOVEMBRE26: "Nov26",
    DECEMBRE26: "Dec26", JANVIER27: "Jan27", FEVRIER27: "Fev27", MARS27: "Mar27"
  };
  function normTxt(s) {
    return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toUpperCase().replace(/[^A-Z0-9 ]+/g, " ").replace(/\s+/g, " ").trim();
  }
  function ventKind(cat) {
    const c = normTxt(cat);
    if (c.indexOf("CARBURANT") >= 0) return "carburant";
    if (c.indexOf("PHARMAC") >= 0) return "pharmacie";
    if (c.indexOf("ALIMENT") >= 0 || c.indexOf("COURSE") >= 0) return "alimentaire";
    return null;
  }
  function monthSheetFor(iso) {
    if (!iso || iso.length < 7) return null;
    const M = { "2026-9": "SEPTEMBRE26", "2026-10": "OCTOBRE26", "2026-11": "NOVEMBRE26",
      "2026-12": "DECEMBRE26", "2027-1": "JANVIER27", "2027-2": "FEVRIER27", "2027-3": "MARS27" };
    return M[parseInt(iso.slice(0, 4), 10) + "-" + parseInt(iso.slice(5, 7), 10)] || null;
  }
  function ventNameFor(sheet, kind) {
    const suf = VENT_MONTHS[sheet];
    if (!suf) return null;
    return (kind === "carburant" ? "VentCarburant" :
            kind === "alimentaire" ? "VentAlimentaire" : "VentPharmacie") + suf;
  }
  /* 18/11 : plages de repli (mêmes coordonnées tous les mois, vérifiées NOV/OCT)
   * — utilisées si la plage nommée est absente (alimentaire Sep/Oct) ou cassée
   * (11 noms posés le 14/11 en constantes String au lieu de plages réelles :
   * santé vérifiée ce jour, cf tests/fo_compta_names_health.py). */
  const VENT_FALLBACK = {
    carburant: "F33:M43",
    alimentaire: "A85:N94",
    pharmacie: "F50:H55"
  };
  function sameDate(v, iso) {
    if (v === null || v === undefined || v === "") return false;
    if (typeof v === "string") return v.slice(0, 10) === iso;
    const n = Number(v);
    if (isNaN(n) || !CM.xlsx || !CM.xlsx.dateSerial) return false;
    const ns = Number(CM.xlsx.dateSerial(iso));
    return !isNaN(ns) && Math.round(n) === Math.round(ns);
  }
  function magasinScore(libelle, header) {
    const l = normTxt(libelle), h = normTxt(header);
    if (!l || !h) return 0;
    if (l === h) return 1000;
    let score = 0;
    h.split(" ").forEach(function (t) { if (t && l.indexOf(t) >= 0) score += t.length; });
    return score;
  }
  function colNum(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }
  function colLetters(n) {
    let s = "";
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  return {
    isConnected: function () {
      const t = loadTokens();
      return !!(t && t.access_token && t.expires_at > Date.now() + 60000);
    },

    hasRefreshToken: function () {
      const t = loadTokens();
      return !!(t && t.refresh_token);
    },

    renew: async function (clientId) {
      const t = loadTokens();
      if (!t || !t.refresh_token) return false;
      const j = await tokenRequest({
        client_id: clientId,
        grant_type: "refresh_token",
        refresh_token: t.refresh_token,
        scope: ["openid", "profile"].concat(CM.getConfig().oauthScopes).join(" ")
      });
      if (!j.refresh_token) j.refresh_token = t.refresh_token;
      j.expires_at = Date.now() + (j.expires_in - 60) * 1000;
      saveTokens(j);
      return true;
    },

    login: async function (clientId, redirectUri, scopes) {
      const st = await pkce();
      sessionStorage.setItem("cm-pkce", JSON.stringify({ verifier: st.verifier }));
      sessionStorage.setItem("cm-oauth-client", clientId);
      const q = new URLSearchParams({
        client_id: clientId,
        response_type: "code",
        redirect_uri: redirectUri,
        response_mode: "fragment",
        scope: ["openid", "profile"].concat(scopes).join(" "),
        code_challenge: st.challenge,
        code_challenge_method: "S256"
      });
      location.href = AUTH + "authorize?" + q.toString();
    },

    handleRedirect: async function () {
      const h = new URLSearchParams(location.hash.slice(1));
      if (!h.get("code")) return false;
      const code = h.get("code");
      const st = JSON.parse(sessionStorage.getItem("cm-pkce") || "null");
      const clientId = sessionStorage.getItem("cm-oauth-client");
      history.replaceState(null, "", location.pathname + location.search);
      if (!st || !clientId) throw new Error("retour OAuth sans état PKCE");
      const t = await tokenRequest({
        client_id: clientId,
        grant_type: "authorization_code",
        code: code,
        redirect_uri: location.origin + location.pathname,
        code_verifier: st.verifier
      });
      t.expires_at = Date.now() + (t.expires_in - 60) * 1000;
      saveTokens(t);
      return true;
    },

    api: async function (path, opts) {
      const t = loadTokens();
      if (!t) throw new Error("non connecté (OAuth) — lancer la connexion dans Diagnostic");
      opts = opts || {};
      opts.headers = Object.assign({
        "Authorization": "Bearer " + t.access_token,
        "Content-Type": "application/json"
      }, opts.headers || {});
      if (_session) opts.headers["workbook-session-id"] = _session;
      const resp = await fetch(path.startsWith("http") ? path : GRAPH + path, opts);
      if (resp.status === 204) return {};
      const txt = await resp.text();
      const j = txt ? JSON.parse(txt) : {};
      if (!resp.ok) throw new Error("Graph " + resp.status + " : " + ((j.error && j.error.message) || resp.statusText));
      return j;
    },

    findWorkbook: async function (path) {
      const p = path.split("/").map(encodeURIComponent).join("/");
      return await this.api("me/drive/root:/" + p + "?select=id,name,size,lastModifiedDateTime");
    },
    findSharedWorkbook: async function (shareUrl) {
      /* Format Graph : "u!" + base64url(URL) — le préfixe u! reste LITTÉRAL,
       * il ne fait PAS partie de la charge encodée (sinon 400 malformed). */
      if (!shareUrl) throw new Error("lien de partage absent (config)");
      const enc = "u!" + b64url(new TextEncoder().encode(shareUrl));
      try {
        return await this.api("shares/" + encodeURIComponent(enc) + "/driveItem");
      } catch (e) {
        throw new Error(e.message + " [lien: " + shareUrl.slice(0, 48) + "...]");
      }
    },
    openSession: async function (itemId, driveId) {
      const j = await this.api(wbBase(itemId, driveId) + "/workbook/createSession",
        { method: "POST", body: JSON.stringify({ persistChanges: true }) });
      _session = j.id;
      _sessionItemId = itemId;
      _sessionDriveId = driveId || null;
    },
    closeSession: async function () {
      if (!_session) return;
      try { await this.api(wbBase(_sessionItemId, _sessionDriveId) + "/workbook/closeSession", { method: "POST" }); } catch (e) { /* tolérant */ }
      _session = null;
      _sessionItemId = null;
      _sessionDriveId = null;
    },
    readJournal: async function (itemId, driveId, sheet, first, last) {
      /* NB : PAS de suffixe /values sur workbookRange — Graph renvoie alors un Json
       * vide ; c'est l'objet range lui-même qui porte "values". (bug trouvé 14/11) */
      const j = await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) +
        "')/range(address='A" + first + ":E" + last + "')");
      return j.values || [];
    },
    writeRow: async function (itemId, driveId, sheet, row, exp) {
      /* Garde anti-écrasement : refuser d'écrire sur une ligne occupée
       * (une écriture de file a écrasé la ligne 8 réelle le 14/11). */
      const cur = await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) +
        "')/range(address='A" + row + ":E" + row + "')");
      const occ = (cur.values || [[]])[0].slice(0, 5).some(function (v) { return v !== null && v !== undefined && v !== ""; });
      if (occ) throw new Error("ligne " + row + " occupée — écriture refusée (anti-écrasement)");
      const values = [[
        CM.xlsx.dateSerial(exp.dateOp),
        CM.xlsx.dateSerial(exp.dateVal || exp.dateOp),
        exp.libelle,
        /* Convention classeur : les depenses sont NEGATIVES en colonne D
         * (Fabrice 15/11, "c'est une depense, ca devrait pas etre en moins ?"). */
        exp.debit === null || exp.debit === undefined ? "" : -Math.abs(exp.debit),
        exp.credit === null || exp.credit === undefined ? "" : exp.credit
      ]];
      await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) +
        "')/range(address='A" + row + ":E" + row + "')",
        { method: "PATCH", body: JSON.stringify({ values: values }) });
    },
    readNamedRange: async function (itemId, driveId, name) {
      return await this.api(wbBase(itemId, driveId) + "/workbook/names('" + encodeURIComponent(name) + "')/range");
    },
    /* Détail de la dépense en commentaire de la cellule cible (Fabrice 18/11).
     * API comments = beta uniquement, SANS session de classeur (valide
     * empiriquement : POST 201 avec cellAddress + contentType "plain", DELETE
     * 204 — tests/fo_compta_comments_probe4.py). Best-effort côté appelant :
     * un échec ne fait pas échouer l'écriture (détail conservé dans l'app). */
    addComment: async function (itemId, driveId, cellAddress, text) {
      const t = loadTokens();
      if (!t || !t.access_token) throw new Error("non connecté (OAuth)");
      const base = driveId ? "drives/" + driveId + "/items/" + itemId : "me/drive/items/" + itemId;
      const resp = await fetch(BETA + base + "/workbook/comments", {
        method: "POST",
        headers: { "Authorization": "Bearer " + t.access_token, "Content-Type": "application/json" },
        body: JSON.stringify({ cellAddress: cellAddress, content: String(text || "").slice(0, 200), contentType: "plain" })
      });
      if (!resp.ok) {
        const j = await resp.json().catch(function () { return {}; });
        throw new Error("commentaire " + resp.status + " : " + ((j.error && j.error.message) || resp.statusText));
      }
      return true;
    },
    /* Ventilation d'une écriture journal dans le tableau du mois (Fabrice 14/11). */
    ventilate: async function (itemId, driveId, exp) {
      const kind = ventKind(exp && exp.categorie);
      if (!kind) return null; /* catégorie sans tableau de ventilation */
      if (!(exp.debit > 0)) return "ventilation " + kind + " : crédit ignoré";
      const sheet = monthSheetFor(exp.dateOp);
      const name = ventNameFor(sheet, kind);
      if (!name) return "ventilation : pas de plage nommée pour " + sheet;
      /* La plage nommée peut manquer (alimentaire Sep/Oct) ou être une constante
       * String cassée (pharmacie tous mois — HTTP 400 InvalidOperation à la
       * lecture) : repli sur les coordonnées littérales du tableau du mois. */
      let rg = null;
      try {
        rg = await this.readNamedRange(itemId, driveId, name);
      } catch (e) {
        rg = null;
      }
      if (!rg || !rg.address) {
        rg = await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) +
          "')/range(address='" + VENT_FALLBACK[kind] + "')");
      }
      const vals = rg.values || [];
      const cells = (rg.address || "").split("!").pop().split("$").join("").split(":");
      let letters = "", digits = "";
      for (let i = 0; i < cells[0].length; i++) {
        if (cells[0][i] >= "0" && cells[0][i] <= "9") digits += cells[0][i]; else letters += cells[0][i];
      }
      const firstRow = parseInt(digits, 10);
      const lastLetters = ((cells[1] || "").replace(/[0-9]/g, "")) || letters;
      const ventWidth = colNum(lastLetters) - colNum(letters) + 1;
      const ventSerial = CM.xlsx && CM.xlsx.dateSerial ? Number(CM.xlsx.dateSerial(exp.dateOp)) : NaN;
      function ventRowEmpty(v) {
        return !(v || []).some(function (x) { return x !== null && x !== undefined && x !== ""; });
      }
      /* Regle Fabrice 14/11 soir : date deja presente -> ligne nouvelle avec la
       * meme date juste en dessous ; date absente -> date la plus proche, ligne
       * nouvelle en dessous avec la date. Jamais de cumul dans une ligne existante. */
      let anchor = -1;
      for (let i = 0; i < vals.length; i++) {
        if (sameDate((vals[i] || [])[0], exp.dateOp)) { anchor = i; break; }
      }
      let rowIdx = -1;
      let scanFrom = 0;
      if (anchor >= 0) {
        scanFrom = anchor + 1;
      } else {
        let best = -1, bestDist = Infinity;
        for (let i = 0; i < vals.length; i++) {
          const dnum = Number((vals[i] || [])[0]);
          if (!isNaN(dnum) && dnum > 0 && !isNaN(ventSerial)) {
            const dist = Math.abs(dnum - ventSerial);
            if (dist < bestDist) { bestDist = dist; best = i; }
          }
        }
        scanFrom = best >= 0 ? best + 1 : 0;
      }
      for (let i = scanFrom; i < vals.length; i++) { if (ventRowEmpty(vals[i])) { rowIdx = i; break; } }
      if (rowIdx < 0) {
        for (let i = 0; i < scanFrom; i++) { if (ventRowEmpty(vals[i])) { rowIdx = i; break; } }
      }
      if (rowIdx < 0) return "ventilation " + kind + " : tableau plein dans " + sheet;
      let colIdx = -1;
      if (kind === "pharmacie") {
        colIdx = 2; /* pharmacie : H = réel */
      } else {
        /* 15/11 : carburant F32:M32 (CARREFOUR + AUTRES inclus), alimentaire
         * A83:N83 ; l'enseigne saisie (menu déroulant) prime, repli libellé. */
        const HDR = kind === "alimentaire" ? "A83:N83" : "F32:M32";
        const hdr = await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) + "')/range(address='" + HDR + "')");
        const names = ((hdr.values || [[""]])[0]) || [];
        let best = 0;
        const target = exp.enseigne || exp.libelle;
        for (let c = 1; c < names.length && c < ventWidth; c++) {
          const s = magasinScore(target, names[c]);
          if (s > best) { best = s; colIdx = c; }
        }
        /* 18/11 : seuil anti-colonne-fausse — un libellé qui ne ressemble
         * qu'à peine à une enseigne (ex « Intermarché » -> INTER.) ne doit
         * pas ventiler dans la mauvaise colonne. */
        if (best < 8) return "ventilation " + kind + " : enseigne non reconnue pour « " + target +
          " » — choisir une enseigne du menu (elle définit la colonne)";
      }
      /* 18/11 : écriture en 2 cellules ciblées (plus de patch de la ligne
       * entière) — un numberFormat ne se pose qu'AVEC values, et « "" » dans
       * un numberFormat de span écraserait le format euro des cellules vides
       * voisines (prouvé ce jour, tests/fo_compta_comments_probe3.py).
       * Format de date affiché : dd/mm/yyyy (choix Fabrice 18/11). */
      if (isNaN(ventSerial)) return "ventilation " + kind + " : date invalide " + exp.dateOp;
      const rowAbs = firstRow + rowIdx;
      const dateCell = colLetters(colNum(letters)) + rowAbs;
      const cell = colLetters(colNum(letters) + colIdx) + rowAbs;
      await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) + "')/range(address='" + dateCell + "')",
        { method: "PATCH", body: JSON.stringify({ values: [[ventSerial]], numberFormat: [["dd/mm/yyyy"]] }) });
      await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) + "')/range(address='" + cell + "')",
        { method: "PATCH", body: JSON.stringify({ values: [[Number(exp.debit)]] }) });
      return { msg: "ventilation " + kind + " -> " + sheet + "!" + cell + " = " + Number(exp.debit) +
        (anchor >= 0 ? " (meme date)" : " (nouvelle date)"), sheet: sheet, cell: cell };
    },
    calculate: async function (itemId, driveId) {
      await this.api(wbBase(itemId, driveId) + "/workbook/application/calculate",
        { method: "POST", body: JSON.stringify({ calculationType: "Full" }) });
    },
    findFreeRowFromValues: function (values, first) {
      for (let i = 0; i < values.length; i++) {
        const r = values[i] || [];
        const isEmpty = r.slice(0, 5).every(function (v) { return v === null || v === undefined || v === ""; });
        if (isEmpty) return first + i;
      }
      return null;
    },
    /* Catégorie -> type de tableau de ventilation (pour la validation saisie). */
    kindFor: ventKind,
    /* Lecture générique d'une plage (en-têtes d'enseignes dynamiques, 18/11). */
    readRange: async function (itemId, driveId, sheet, addr) {
      const j = await this.api(wbBase(itemId, driveId) + "/workbook/worksheets('" + encodeURIComponent(sheet) +
        "')/range(address='" + addr + "')");
      return j.values || [];
    },
    /* Mois -> feuille (peuplement dynamique des enseignes selon la date). */
    monthSheetFor: monthSheetFor
  };
})();