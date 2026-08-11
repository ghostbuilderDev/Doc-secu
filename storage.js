(() => {
  "use strict";

  const DB_NAME = "docu-secu-isf";
  const DB_VERSION = 1;
  const STORES = ["isfs", "chantiers", "library", "settings"];
  const FALLBACK_KEY = "docu-secu-isf-fallback-v1";
  let dbPromise = null;
  let fallback = false;

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve) => {
      if (!("indexedDB" in window)) {
        fallback = true;
        resolve(null);
        return;
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        STORES.forEach((name) => {
          if (!db.objectStoreNames.contains(name)) db.createObjectStore(name, { keyPath: "id" });
        });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => {
        console.warn("IndexedDB indisponible, utilisation du stockage local.", request.error);
        fallback = true;
        resolve(null);
      };
    });
    return dbPromise;
  }

  async function getAll(storeName) {
    validateStore(storeName);
    const db = await open();
    if (!db || fallback) return readFallback()[storeName] || [];
    return requestResult(db.transaction(storeName, "readonly").objectStore(storeName).getAll());
  }

  async function get(storeName, id) {
    validateStore(storeName);
    const db = await open();
    if (!db || fallback) return (readFallback()[storeName] || []).find((item) => item.id === id) || null;
    return requestResult(db.transaction(storeName, "readonly").objectStore(storeName).get(id));
  }

  async function put(storeName, value) {
    validateStore(storeName);
    if (!value || typeof value !== "object" || !value.id) throw new Error(`Identifiant manquant pour ${storeName}.`);
    const copy = clone(value);
    const db = await open();
    if (!db || fallback) {
      const data = readFallback();
      const items = data[storeName] || [];
      const index = items.findIndex((item) => item.id === copy.id);
      if (index >= 0) items[index] = copy;
      else items.push(copy);
      data[storeName] = items;
      writeFallback(data);
      return copy;
    }
    await requestResult(db.transaction(storeName, "readwrite").objectStore(storeName).put(copy));
    return copy;
  }

  async function remove(storeName, id) {
    validateStore(storeName);
    const db = await open();
    if (!db || fallback) {
      const data = readFallback();
      data[storeName] = (data[storeName] || []).filter((item) => item.id !== id);
      writeFallback(data);
      return;
    }
    await requestResult(db.transaction(storeName, "readwrite").objectStore(storeName).delete(id));
  }

  async function clear(storeName) {
    validateStore(storeName);
    const db = await open();
    if (!db || fallback) {
      const data = readFallback();
      data[storeName] = [];
      writeFallback(data);
      return;
    }
    await requestResult(db.transaction(storeName, "readwrite").objectStore(storeName).clear());
  }

  async function exportAll() {
    const payload = {
      schema: "docu-secu-isf-backup",
      version: 1,
      exportedAt: new Date().toISOString(),
      data: {},
    };
    for (const store of STORES) payload.data[store] = await getAll(store);
    return payload;
  }

  async function importAll(payload, options = {}) {
    if (!payload || payload.schema !== "docu-secu-isf-backup" || !payload.data) throw new Error("Ce fichier n’est pas une sauvegarde Docu Sécurité valide.");
    const replace = Boolean(options.replace);
    if (replace) {
      for (const store of STORES) await clear(store);
    }
    const counts = {};
    for (const store of STORES) {
      const items = Array.isArray(payload.data[store]) ? payload.data[store] : [];
      for (const item of items) {
        if (item?.id) await put(store, item);
      }
      counts[store] = items.length;
    }
    return counts;
  }

  function requestResult(request) {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Erreur IndexedDB."));
    });
  }

  function validateStore(name) {
    if (!STORES.includes(name)) throw new Error(`Stockage inconnu : ${name}`);
  }

  function readFallback() {
    try {
      const value = JSON.parse(localStorage.getItem(FALLBACK_KEY) || "{}");
      return STORES.reduce((result, name) => ({ ...result, [name]: Array.isArray(value[name]) ? value[name] : [] }), {});
    } catch {
      return STORES.reduce((result, name) => ({ ...result, [name]: [] }), {});
    }
  }

  function writeFallback(data) {
    localStorage.setItem(FALLBACK_KEY, JSON.stringify(data));
  }

  function clone(value) {
    return typeof structuredClone === "function" ? structuredClone(value) : JSON.parse(JSON.stringify(value));
  }

  window.ISFStorage = { open, getAll, get, put, remove, clear, exportAll, importAll, stores: [...STORES] };
})();
