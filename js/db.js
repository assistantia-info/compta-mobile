/* Compta Mobile — file locale IndexedDB (offline-first)
 * Store « depenses » : une ligne de journal en attente (pending),
 * poussée (pushed) ou en erreur (error après 5 essais).
 * uuid = clé anti-doublon générée à la saisie, conservée jusqu'au push.
 */
"use strict";
window.CM = window.CM || {};

CM.db = (function () {
  const DB_NAME = "compta-mobile";
  const DB_VERSION = 1;
  let _db = null;

  function open() {
    if (_db) return Promise.resolve(_db);
    return new Promise(function (resolve, reject) {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function (ev) {
        const db = ev.target.result;
        if (!db.objectStoreNames.contains("depenses")) {
          const st = db.createObjectStore("depenses", { keyPath: "id", autoIncrement: true });
          st.createIndex("by_uuid", "uuid", { unique: true });
          st.createIndex("by_status", "status", { unique: false });
        }
        if (!db.objectStoreNames.contains("meta")) {
          db.createObjectStore("meta", { keyPath: "key" });
        }
      };
      req.onsuccess = function () { _db = req.result; resolve(_db); };
      req.onerror = function () { reject(req.error); };
    });
  }

  function req2p(req) {
    return new Promise(function (resolve, reject) {
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }

  function withStore(name, mode, fn) {
    return open().then(function (db) {
      return new Promise(function (resolve, reject) {
        const t = db.transaction(name, mode);
        let result;
        Promise.resolve(fn(t.objectStore(name))).then(function (r) { result = r; }, reject);
        t.oncomplete = function () { resolve(result); };
        t.onerror = t.onabort = function () { reject(t.error); };
      });
    });
  }

  return {
    addDepense: function (rec) {
      rec.uuid = rec.uuid || (crypto.randomUUID ? crypto.randomUUID()
                         : "u-" + Date.now() + "-" + Math.random().toString(36).slice(2));
      rec.status = rec.status || "pending";
      rec.attempts = rec.attempts || 0;
      rec.row = rec.row || null;
      rec.lastError = rec.lastError || null;
      rec.createdAt = rec.createdAt || new Date().toISOString();
      rec.updatedAt = new Date().toISOString();
      return withStore("depenses", "readwrite", function (s) { return req2p(s.add(rec)); })
        .then(function () { return rec; });
    },
    updateDepense: function (rec) {
      rec.updatedAt = new Date().toISOString();
      return withStore("depenses", "readwrite", function (s) { return req2p(s.put(rec)); })
        .then(function () { return rec; });
    },
    deleteDepense: function (id) {
      return withStore("depenses", "readwrite", function (s) { return req2p(s.delete(id)); });
    },
    listDepenses: function () {
      return withStore("depenses", "readonly", function (s) { return req2p(s.getAll()); });
    },
    getMeta: function (key) {
      return withStore("meta", "readonly", function (s) { return req2p(s.get(key)); })
        .then(function (r) { return r ? r.value : null; });
    },
    setMeta: function (key, value) {
      return withStore("meta", "readwrite", function (s) { return req2p(s.put({ key: key, value: value })); });
    }
  };
})();