/* Compta Mobile — logique UI (V1)
 * Onglets : Saisie / File / Réglages / Diagnostic.
 * Toutes les écritures passent par CM.sync (dry-run par défaut).
 * L'onglet Diagnostic exécute les tests de reprise (a) OAuth et (b) lien
 * de partage documentés dans le readme projet (section PAUSE).
 */
"use strict";
(function () {
  const $ = function (id) { return document.getElementById(id); };
  const ui = { type: "debit" };
  /* 15/11 : enseignes normalisées (en-têtes réels des tables, carburant
   * F32:M32, alimentaire A83:N83). Menu déroulant à la saisie ; la
   * ventilation matche par nom normalisé, repli sur le libellé.
   * 18/11 : cette liste devient le REPLI statique — les menus sont peuplés
   * dynamiquement depuis les en-têtes du tableur pour le mois de la date
   * saisie (cache localStorage, hors ligne = repli). */
  const ENSEIGNES = {
    carburant: ["LECLERC MENDES", "PETIT LECLERC", "INTER. BEAUVOIR", "INTER. AIFFRES", "CARREFOUR", "AUTRES 1", "AUTRES 2"],
    alimentaire: ["LECLERC DRIVE BESSINES", "PETIT LECLERC DRIVE", "LECLERC DRIVE MENDES", "LECLERC", "LIDL", "INTER. BEAUVOIR", "INTER. AIFFRES drive", "PIZZA/NUGGETS/kebab", "Action/Amazon/PointVert", "boucherie", "boulanger", "Jardin de la Noue / Saule au Jardin", "netto - Clé des champs"]
  };
  function normCat(s) {
    const c = String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
    if (c.indexOf("CARBURANT") >= 0) return "carburant";
    if (c.indexOf("ALIMENT") >= 0 || c.indexOf("COURSE") >= 0) return "alimentaire";
    if (c.indexOf("PHARMAC") >= 0) return "pharmacie";
    return null;
  }
  /* 18/11 : menus peuplés depuis le tableur. Catégorie = select strict
   * (plus de saisie libre -> plus de catégorie sans table). Enseignes =
   * en-têtes réels du mois de la date saisie, lus via Graph (lecture seule),
   * mis en cache en localStorage (« cm-headers ») et servis depuis le cache
   * hors ligne ; repli = liste statique ENSEIGNES. */
  const HDR_KEY = "cm-headers";
  let HDR_CACHE = {};
  try { HDR_CACHE = JSON.parse(localStorage.getItem(HDR_KEY) || "{}") || {}; } catch (e) { HDR_CACHE = {}; }
  let hdrBusy = false;
  function cleanHeader(s) { return String(s || "").replace(/\s+/g, " ").trim(); }
  function monthSheet() { return CM.graph.monthSheetFor($("f-date").value || todayISO()); }
  function fillEnseigne() {
    const sel = $("f-enseigne");
    if (!sel) return;
    const kind = normCat($("f-categorie").value);
    const dyn = monthSheet() ? HDR_CACHE[monthSheet()] : null;
    let list = (dyn && dyn[kind] && dyn[kind].length) ? dyn[kind] : (ENSEIGNES[kind] || []);
    if (kind === "pharmacie") list = []; /* colonne fixe (H = réel), pas de menu */
    sel.innerHTML = "";
    const o0 = document.createElement("option");
    o0.value = "";
    o0.textContent = !kind ? "— choisir la catégorie d'abord —"
      : (list.length ? "— choisir un magasin —" : "— sans objet pour cette catégorie —");
    sel.appendChild(o0);
    list.forEach(function (n) {
      const o = document.createElement("option");
      o.value = n; o.textContent = n;
      sel.appendChild(o);
    });
  }
  async function refreshEnseignes() {
    const cfg = CM.getConfig();
    if (cfg.authMode !== "oauth" || !CM.graph.isConnected()) return;
    const sheet = monthSheet();
    if (!sheet || HDR_CACHE[sheet] || hdrBusy) return;
    hdrBusy = true;
    try {
      const item = cfg.workbookMode === "share" && cfg.workbookShareUrl
        ? await CM.graph.findSharedWorkbook(cfg.workbookShareUrl)
        : await CM.graph.findWorkbook(cfg.workbookPath);
      const driveId = (item.parentReference && item.parentReference.driveId) || null;
      const car = await CM.graph.readRange(item.id, driveId, sheet, "F32:M32");
      const ali = await CM.graph.readRange(item.id, driveId, sheet, "A83:N83");
      HDR_CACHE[sheet] = {
        carburant: (car[0] || []).slice(1).map(cleanHeader).filter(Boolean),
        alimentaire: (ali[0] || []).slice(1).map(cleanHeader).filter(Boolean)
      };
      localStorage.setItem(HDR_KEY, JSON.stringify(HDR_CACHE));
    } catch (e) {
      console.warn("enseignes dynamiques :", e); /* repli statique, saisie non bloquée */
    }
    hdrBusy = false;
    fillEnseigne();
    const cur = monthSheet();
    if (cur && cur !== sheet && !HDR_CACHE[cur]) refreshEnseignes();
  }

  function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function fmtDate(iso) { return String(iso || "").slice(0, 10).split("-").reverse().join("/"); }
  function fmtMontant(r) {
    const d = r.debit !== null && r.debit !== undefined;
    return (d ? "−" : "+") + (d ? r.debit : r.credit) + " €";
  }
  function badge(rec) {
    if (rec.status === "pushed") return '<span class="badge badge-ok">' + (rec.row ? "poussé · ligne " + rec.row : "ventilé") + "</span>";
    if (rec.status === "error") return '<span class="badge badge-err">erreur</span>';
    if (rec.simulatedAt) return '<span class="badge badge-dry">' + (rec.row ? "simulé · ligne " + rec.row : "simulé") + "</span>";
    return '<span class="badge badge-wait">en attente' + (rec.attempts > 0 ? " · essai " + rec.attempts : "") + "</span>";
  }
  function todayISO() {
    const d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }

  async function refresh() {
    const all = (await CM.db.listDepenses()).sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
    const pending = all.filter(function (r) { return r.status === "pending"; }).length;
    $("pending-badge").textContent = pending + " en file";
    const html = all.slice(0, 30).map(function (r) {
      return "<li><div class='li-main'><span class='li-date'>" + fmtDate(r.dateOp) + "</span><span>" +
        esc(r.enseigne || r.libelle || "?") + "</span><span class='li-amt'>" + fmtMontant(r) + "</span></div><div class='li-sub'>" +
        (r.categorie ? "<span class='badge'>" + esc(r.categorie) + "</span> " : "") +
        (r.libelle ? "<span class='muted'>" + esc(r.libelle) + "</span> " : "") +
        badge(r) + (r.lastError ? " — <span class='err'>" + esc(r.lastError) + "</span>" : "") +
        /* 18/11 : Supprimer sur TOUT item non pousse (les bloques restent au
         * statut pending tant que le cap 5 essais n'est pas atteint — les
         * boutons ne s'affichaient donc pas, constat Fabrice v18) ;
         * Reessayer des qu'il y a eu une tentative ou une erreur. */
        (r.status !== "pushed" ? (r.attempts > 0 || r.lastError ? " <button class='btn-mini' data-act='retry' data-id='" + r.id + "'>Réessayer</button>" : "") +
        " <button class='btn-mini' data-act='del' data-id='" + r.id + "'>Supprimer</button>" : "") + "</div></li>";
    }).join("");
    $("recent-list").innerHTML = html || '<li class="empty">Rien pour le moment.</li>';
    $("file-list").innerHTML = html || '<li class="empty">File vide.</li>';
  }

  /* 18/11 : actions Réessayer / Supprimer — DELEGATION d'événements attachée
   * une seule fois en init (les listeners par rendu mouraient à chaque
   * re-render de liste, et une exception dans un handler async restait
   * invisible : constat Fabrice v19 « clic mais rien ne s'efface »).
   * Toute erreur est maintenant AFFICHÉE (alert) au lieu d'être avalée. */
  async function handleFileAction(act, idStr) {
    try {
      const all = await CM.db.listDepenses();
      const rec = all.find(function (x) { return x.id === Number(idStr); });
      if (!rec) { alert("Item introuvable (id " + idStr + ") — fermez et rouvrez l'app puis réessayez."); return; }
      if (act === "retry") {
        rec.status = "pending"; rec.attempts = 0; rec.nextAttemptAt = null; rec.lastError = null;
        await CM.db.updateDepense(rec);
        CM.sync.run("réessai");
      } else if (act === "del") {
        await CM.db.deleteDepense(rec.id);
        await refresh();
      }
    } catch (e) {
      alert("Action « " + act + " » échouée : " + ((e && e.message) || e));
      await refresh();
    }
  }
  function wireFileActions() {
    ["file-list", "recent-list"].forEach(function (id) {
      const el = $(id);
      if (!el || el.dataset.wired) return;
      el.dataset.wired = "1";
      el.addEventListener("click", function (ev) {
        const b = ev.target && ev.target.closest ? ev.target.closest(".btn-mini") : null;
        if (!b) return;
        handleFileAction(b.dataset.act, b.dataset.id);
      });
    });
  }

  function tabActivate(name) {
    document.querySelectorAll(".tab").forEach(function (b) { b.classList.toggle("active", b.dataset.tab === name); });
    document.querySelectorAll(".panel").forEach(function (p) { p.classList.toggle("active", p.id === "tab-" + name); });
  }

  async function submitDepense(ev) {
    ev.preventDefault();
    const msg = $("saisie-msg");
    const montant = parseFloat($("f-montant").value.replace(",", "."));
    const libelle = $("f-libelle").value.trim();
    const categorie = $("f-categorie").value.trim();
    if (isNaN(montant) || montant <= 0) { msg.hidden = false; msg.textContent = "Montant invalide."; return; }
    if (!categorie) { msg.hidden = false; msg.textContent = "Catégorie obligatoire (Carburant, Alimentation ou Pharmacie)."; return; }
    /* 18/11 : le détail devient le champ « libellé » (commentaire en cellule
     * cible), le magasin est l'enseigne du menu déroulant. Une catégorie sans
     * table de ventilation (ex Divers) échouerait à chaque push — refusée
     * dès la saisie. */
    if (!CM.graph.kindFor(categorie)) {
      msg.hidden = false;
      msg.textContent = "Catégorie « " + categorie + " » sans table de ventilation (Carburant, Alimentation ou Pharmacie uniquement). Saisie non enregistrée.";
      return;
    }
    /* 18/11 : le magasin est obligatoire (hors pharmacie) — c'est lui qui
     * définit la colonne du tableur ; sans lui, la ventilation refuserait. */
    const enseigneVal = $("f-enseigne") ? ($("f-enseigne").value || null) : null;
    if (CM.graph.kindFor(categorie) !== "pharmacie" && !enseigneVal) {
      msg.hidden = false;
      msg.textContent = "Magasin obligatoire (il définit la colonne du tableur).";
      return;
    }
    const rec = {
      dateOp: $("f-date").value || todayISO(),
      dateVal: $("f-date").value || todayISO(),
      libelle: libelle,
      categorie: categorie,
      enseigne: enseigneVal,
      debit: ui.type === "debit" ? montant : null,
      credit: ui.type === "credit" ? montant : null
    };
    await CM.db.addDepense(rec);
    msg.hidden = false;
    msg.textContent = "Enregistré — " + (CM.getConfig().dryRun
      ? "mode SIMULATION : rien n'est envoyé vers OneDrive."
      : "la synchronisation vers OneDrive va démarrer.");
    $("f-libelle").value = "";
    $("f-categorie").value = "";
    if ($("f-enseigne")) $("f-enseigne").value = "";
    $("f-montant").value = "";
    $("f-libelle").focus();
    await refresh();
    CM.sync.run("saisie");
  }

  function fillSettings() {
    const c = CM.getConfig();
    $("c-dryrun").checked = c.dryRun;
    $("c-authmode").value = c.authMode;
    $("c-shareurl").value = c.shareUrl;
    $("c-clientid").value = c.oauthClientId;
    $("c-path").value = c.workbookPath;
    $("c-sheet").value = c.journalSheet;
    $("hint-redirect").textContent = location.origin + location.pathname;
    updateGroups();
  }
  function updateGroups() {
    const mode = $("c-authmode").value;
    $("grp-link").style.display = mode === "link" ? "" : "none";
    $("grp-oauth").style.display = mode === "oauth" ? "" : "none";
  }
  function saveSettings(ev) {
    ev.preventDefault();
    const c = CM.getConfig();
    c.dryRun = $("c-dryrun").checked;
    c.authMode = $("c-authmode").value;
    c.shareUrl = $("c-shareurl").value.trim();
    c.oauthClientId = $("c-clientid").value.trim();
    c.workbookPath = $("c-path").value.trim();
    c.journalSheet = $("c-sheet").value.trim();
    CM.saveConfig(c);
    renderModeBadge();
    alert("Réglages enregistrés.");
  }
  function renderModeBadge() {
    const c = CM.getConfig();
    const b = $("mode-badge");
    b.textContent = c.dryRun ? "SIMULATION" : "LIVE";
    b.className = "badge " + (c.dryRun ? "badge-dry" : "badge-live");
  }
  function netBadge() {
    $("net-badge").textContent = navigator.onLine ? "en ligne" : "hors ligne";
    $("net-badge").className = "badge " + (navigator.onLine ? "badge-ok" : "badge-err");
  }
  /* 18/11 : badge connexion Graph toujours visible en tete — avant, un jeton
   * expire etait invisible (constat Fabrice : « tout restait en attente »
   * jusqu'a un test manuel en Diagnostic). */
  function renderAuthBadge() {
    const b = $("auth-badge");
    if (!b) return;
    const cfg = CM.getConfig();
    if (cfg.authMode !== "oauth") { b.textContent = "lien"; b.className = "badge"; return; }
    const ok = CM.graph.isConnected();
    b.textContent = ok ? "Graph OK" : "Graph off";
    b.className = "badge " + (ok ? "badge-ok" : "badge-err");
  }

  function stepsText(r) {
    return r.steps.map(function (s) { return (s.ok ? "OK   " : "KO   ") + s.step + " : " + s.detail; }).join("\n");
  }
  async function diagGet() {
    const out = $("diag-out");
    out.textContent = "… test GET en cours …\n";
    try {
      const cfg = CM.getConfig();
      if (!cfg.shareUrl) throw new Error("Renseigner le lien de partage dans Réglages d'abord.");
      const r = await CM.link.diagnose(cfg.shareUrl, {});
      out.textContent = stepsText(r) + "\n\nVERDICT lecture : " +
        (r.journal && r.journal.firstFreeRow
          ? "RÉUSSI — critère (b) lecture OK, première ligne libre = " + r.journal.firstFreeRow
          : "à examiner (journal plein ou illisible)");
    } catch (e) {
      out.textContent = "ÉCHEC : " + (e.message || e) +
        "\n\nSi l'erreur est de type CORS ou réseau, voir README-PWA.md (section Limites) : " +
        "le test doit alors être lancé depuis un contexte autorisé (http://localhost ou GitHub Pages).";
    }
  }
  async function diagPut() {
    if (!confirm("Test d'écriture : le classeur sera ré-uploadé à l'octet identique (seule la date de modification changera). Continuer ?")) return;
    const out = $("diag-out");
    out.textContent = "… test GET + PUT en cours …\n";
    try {
      const cfg = CM.getConfig();
      const r = await CM.link.diagnose(cfg.shareUrl, { putTest: true });
      out.textContent = stepsText(r) + "\n\nVERDICT écriture : " +
        (r.putOk ? "OK — critère (b) VALIDÉ, la voie lien de partage est utilisable (désactiver dry-run pour pousser réellement)."
                 : "non validé");
    } catch (e) {
      out.textContent = "ÉCHEC : " + (e.message || e) + "\n(lecture = officiellement supportée ; écriture PUT = non garantie par Microsoft — voir readme projet, section PAUSE)";
    }
  }
  async function diagOauth() {
    const cfg = CM.getConfig();
    if (!CM.graph.isConnected()) {
      if (!cfg.oauthClientId) { alert("Renseigner le client ID dans Réglages d'abord."); return; }
      await CM.graph.login(cfg.oauthClientId, location.origin + location.pathname, cfg.oauthScopes);
      return; /* redirection Microsoft — au retour, handleRedirect() connectera */
    }
    const out = $("diag-out");
    out.textContent = "… test Graph en cours …\n";
    try {
      const item = await CM.graph.findWorkbook(cfg.workbookPath);
      const vals = await CM.graph.readJournal(item.id, (item.parentReference && item.parentReference.driveId) || null,
        cfg.journalSheet, cfg.journalFirstRow, cfg.journalLastRow);
      const row = CM.graph.findFreeRowFromValues(vals, cfg.journalFirstRow);
      out.textContent = "OK classeur : " + item.name + " (" + item.size + " octets)\n" +
        "OK journal " + cfg.journalSheet + " : " + vals.length + " lignes lues, première libre = " + row +
        "\n\nVERDICT : critère (a) lecture OK — connexion + classeur + journal accessibles.";
    } catch (e) {
      out.textContent = "ÉCHEC : " + (e.message || e);
    }
  }

  async function init() {
    try { await CM.graph.handleRedirect(); } catch (e) { console.warn("OAuth redirect :", e); }
    if (!CM.graph.isConnected() && CM.graph.hasRefreshToken()) {
      try { await CM.graph.renew(CM.getConfig().oauthClientId); }
      catch (e) { console.warn("reconnexion silencieuse :", e); }
    }
    $("f-date").value = todayISO();
    fillSettings();
    renderModeBadge();
    netBadge();
    renderAuthBadge();
    refresh();
    document.querySelectorAll(".tab").forEach(function (b) {
      b.addEventListener("click", function () { tabActivate(b.dataset.tab); });
    });
    document.querySelectorAll("#seg-type .seg-btn").forEach(function (b) {
      b.addEventListener("click", function () {
        ui.type = b.dataset.val;
        document.querySelectorAll("#seg-type .seg-btn").forEach(function (x) { x.classList.toggle("active", x === b); });
      });
    });
    $("form-depense").addEventListener("submit", submitDepense);
    $("f-categorie").addEventListener("input", fillEnseigne);
    $("f-date").addEventListener("change", function () { fillEnseigne(); refreshEnseignes(); });
    fillEnseigne();
    refreshEnseignes();
    wireFileActions();
    $("form-config").addEventListener("submit", saveSettings);
    $("c-authmode").addEventListener("change", updateGroups);
    $("btn-sync").addEventListener("click", function () { CM.sync.run("manuel"); });
    /* 18/11 : indicateur de version visible (absent avant — Fabrice ne pouvait
     * pas verifier qu'il avait la derniere app). */
    (function () {
      const el = $("diag-version");
      if (!el) return;
      let swState = "sans service worker";
      if ("serviceWorker" in navigator && navigator.serviceWorker.controller) {
        swState = "service worker actif";
        if (navigator.serviceWorker.ready) {
          navigator.serviceWorker.ready.then(function (reg) { return reg.active ? reg.active.scriptURL : ""; })
            .catch(function () {});
        }
      }
      let cachesTxt = "";
      if (window.caches && caches.keys) {
        caches.keys().then(function (keys) {
          cachesTxt = keys.length ? " · cache " + keys.join(", ") : "";
          el.textContent = "Version : " + CM.APP_VERSION + " · " + swState + cachesTxt;
        }).catch(function () {
          el.textContent = "Version : " + CM.APP_VERSION + " · " + swState;
        });
      } else {
        el.textContent = "Version : " + CM.APP_VERSION + " · " + swState;
      }
    })();
    $("btn-diag-get").addEventListener("click", diagGet);
    $("btn-diag-put").addEventListener("click", diagPut);
    $("btn-diag-oauth").addEventListener("click", diagOauth);
    CM.sync.on(function (ev) {
      if (ev.type === "start" || ev.type === "done") renderAuthBadge();
      if (ev.type === "log") {
        const el = $("sync-log");
        el.textContent += "[" + ev.at.slice(11, 19) + "] " + ev.msg + "\n";
        el.scrollTop = el.scrollHeight;
      }
      if (ev.type === "done") refresh();
    });
    window.addEventListener("online", function () { netBadge(); CM.sync.run("online"); });
    window.addEventListener("offline", netBadge);
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("sw.js").catch(function (e) { console.warn("sw :", e); });
    }
    CM.sync.run("démarrage");
  }

  document.addEventListener("DOMContentLoaded", init);
})();