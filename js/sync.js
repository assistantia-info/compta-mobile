/* Compta Mobile — moteur de synchronisation
 * FIFO sur la file IndexedDB. Allocation de ligne par LECTURE du classeur
 * (source de vérité = le journal, pas un curseur local — robuste aux saisies
 * manuelles de Fabrice dans Excel). Anti-doublon : uuid + idempotence par
 * ligne (link : re-PUT complet déterministe ; oauth : PATCH même contenu).
 * Retries : backoff 30 s / 2 min / 10 min, cap 5 essais puis statut error.
 * dryRun=true : poussée SIMULÉE, aucune requête d'écriture n'est émise.
 */
"use strict";
window.CM = window.CM || {};
CM.sync = (function () {
  const BACKOFF_MS = [30000, 120000, 600000];
  let running = false;
  const listeners = [];

  function emit(ev) { listeners.forEach(function (fn) { try { fn(ev); } catch (e) {} }); }
  function log(msg) { emit({ type: "log", msg: msg, at: new Date().toISOString() }); }

  async function getAdapter(cfg) {
    if (cfg.authMode === "oauth") {
      if (!cfg.oauthClientId) throw new Error("mode OAuth sans client ID (onglet Réglages)");
      /* 18/11 : renouvellement automatique AU MOMENT de la synchro — avant,
       * un jeton expire laissait tout en attente jusqu'a un test manuel en
       * Diagnostic (constat Fabrice). Echec de renouvellement -> message
       * actionnable clair dans le journal. */
      if (!CM.graph.isConnected()) {
        if (CM.graph.hasRefreshToken()) {
          try {
            await CM.graph.renew(cfg.oauthClientId);
          } catch (e) {
            log("renouvellement jeton impossible : " + (e.message || e));
          }
        }
        if (!CM.graph.isConnected()) {
          throw new Error("OAuth non connecté — Diagnostic → « Tester connexion Microsoft Graph » pour vous reconnecter");
        }
      }
      const item = cfg.workbookMode === "share" && cfg.workbookShareUrl
        ? await CM.graph.findSharedWorkbook(cfg.workbookShareUrl)
        : await CM.graph.findWorkbook(cfg.workbookPath);
      item.driveId = item.parentReference && item.parentReference.driveId;
      await CM.graph.openSession(item.id, item.driveId);
      return { kind: "oauth", item: item };
    }
    if (cfg.authMode === "link") {
      if (!cfg.shareUrl) throw new Error("mode lien de partage sans URL (onglet Réglages)");
      return { kind: "link" };
    }
    throw new Error("mode inconnu : " + cfg.authMode);
  }

  async function findFreeRow(cfg, adapter) {
    if (adapter.kind === "oauth") {
      const vals = await CM.graph.readJournal(adapter.item.id, adapter.item.driveId, cfg.journalSheet, cfg.journalFirstRow, cfg.journalLastRow);
      return CM.graph.findFreeRowFromValues(vals, cfg.journalFirstRow);
    }
    const st = await CM.link.readJournalState(cfg.shareUrl, cfg.journalSheet, cfg.journalFirstRow, cfg.journalLastRow);
    return st.firstFreeRow;
  }

  async function runSync(reason) {
    if (running) return { skipped: true };
    running = true;
    emit({ type: "start", reason: reason || "?" });
    const cfg = CM.getConfig();
    let pushed = 0, failed = 0, simulated = 0;
    let adapter = null;
    try {
      adapter = await getAdapter(cfg);
      /* 15/11 : voie OAuth = ventilation directe dans la table du mois, plus
       * d'écriture journal (les onglets REL sont la copie bancaire). */
      let simRow = null;
      if (adapter.kind !== "oauth") {
        simRow = await findFreeRow(cfg, adapter);
        if (simRow === null) throw new Error("journal plein (lignes " + cfg.journalFirstRow + "-" + cfg.journalLastRow + ")");
      }
      const all = await CM.db.listDepenses();
      const pending = all.filter(function (r) { return r.status === "pending"; })
                         .sort(function (a, b) { return a.createdAt < b.createdAt ? -1 : 1; });
      for (const rec of pending) {
        if (rec.attempts > 0 && rec.nextAttemptAt && Date.now() < rec.nextAttemptAt) continue; /* backoff en cours */
        if (cfg.dryRun) {
          if (adapter.kind === "oauth") {
            log("[SIMUL] ventilation " + (rec.categorie || "?") + " : " + rec.dateOp + " | " +
                rec.libelle + " | enseigne " + (rec.enseigne || "-") + " | " +
                (rec.debit !== null && rec.debit !== undefined ? "débit " + rec.debit : "crédit " + rec.credit));
          } else {
            if (simRow === null) { log("journal plein — simulation arrêtée"); break; }
            log("[SIMUL] ligne " + simRow + " : " + rec.dateOp + " | " + rec.libelle + " | " +
                (rec.debit !== null && rec.debit !== undefined ? "débit " + rec.debit : "crédit " + rec.credit));
            rec.row = simRow;
            simRow++;
          }
          rec.simulatedAt = new Date().toISOString();
          await CM.db.updateDepense(rec);
          simulated++;
          continue;
        }
        try {
          if (adapter.kind === "oauth") {
            /* 15/11 : l'entrée s'écrit directement dans la table du mois ; un
             * échec de ventilation est un échec d'entrée (retry avec backoff). */
            log("push (Graph) ventilation : " + rec.libelle);
            const v = await CM.graph.ventilate(adapter.item.id, adapter.item.driveId, rec);
            const vmsg = v && v.msg ? v.msg : v;
            if (!vmsg) throw new Error("ventilation : catégorie sans tableau (ni carburant ni alimentaire ni pharmacie)");
            if (vmsg.indexOf(" -> ") < 0) throw new Error(vmsg); /* enseigne non reconnue etc. */
            log(vmsg);
            /* 18/11 : détail de la dépense en commentaire de la cellule cible
             * (best-effort — un échec ne remet pas l'écriture en question). */
            if (v && v.sheet && v.cell && rec.libelle) {
              try {
                await CM.graph.addComment(adapter.item.id, adapter.item.driveId || null, v.sheet + "!" + v.cell, rec.libelle);
                log("commentaire ajouté : " + rec.libelle);
              } catch (ce) {
                log("commentaire impossible (détail conservé dans l'app) : " + (ce.message || ce));
              }
            }
            rec.row = null;
          } else {
            log("push (lien) : " + rec.libelle);
            const r = await CM.link.pushExpense(cfg.shareUrl, rec, cfg.journalSheet, cfg.journalFirstRow, cfg.journalLastRow);
            rec.row = r.row;
          }
          rec.status = "pushed";
          rec.pushedAt = new Date().toISOString();
          rec.lastError = null;
          await CM.db.updateDepense(rec);
          log("OK : " + rec.libelle + " (ventilation effectuée)");
          pushed++;
        } catch (e) {
          const emsg = String(e.message || e);
          /* 18/11 : erreurs permanentes (catégorie/enseigne/date sans destination)
           * — aucun retry ne les résout, statut error immédiat (bouton Réessayer
           * dans l'onglet File après correction de la saisie). */
          if (/enseigne non reconnue|catégorie sans tableau|tableau plein|pas de plage nommée|date invalide/.test(emsg)) {
            rec.status = "error";
            rec.attempts = 5;
            rec.lastError = emsg;
            rec.nextAttemptAt = null;
            await CM.db.updateDepense(rec);
            log("ERREUR DEFINITIVE : " + emsg);
            failed++;
            continue;
          }
          rec.attempts++;
          rec.lastError = String(e.message || e);
          const wait = BACKOFF_MS[Math.min(rec.attempts - 1, BACKOFF_MS.length - 1)] || 600000;
          rec.nextAttemptAt = Date.now() + wait;
          await CM.db.updateDepense(rec);
          log("ÉCHEC (essai " + rec.attempts + ", retry dans " + Math.round(wait / 1000) + " s) : " + rec.lastError);
          failed++;
          if (rec.attempts >= 5) {
            rec.status = "error";
            await CM.db.updateDepense(rec);
            log("abandon après 5 essais : " + rec.libelle + " (statut error, retry manuel via nouvelle saisie)");
          }
        }
      }
      if (adapter.kind === "oauth" && pushed > 0) {
        await CM.graph.calculate(adapter.item.id, adapter.item.driveId); /* propage la chaîne mensuelle une fois par run */
        log("recalcul classeur (calculate Full) effectué");
      }
    } catch (e) {
      log("SYNC INTERROMPUE : " + (e.message || e));
      emit({ type: "error", error: String(e.message || e) });
    } finally {
      if (adapter && adapter.kind === "oauth") await CM.graph.closeSession();
      running = false;
      emit({ type: "done", pushed: pushed, failed: failed, simulated: simulated });
    }
    return { pushed: pushed, failed: failed, simulated: simulated };
  }

  return {
    run: runSync,
    isRunning: function () { return running; },
    on: function (fn) { listeners.push(fn); }
  };
})();