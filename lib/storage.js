/**
 * IndexedDB storage layer for EtiquetaFit Dashboard
 * Stores labels, custom categories, active checkout cart, and cached PDFs.
 */

const DB_NAME = 'etiquetafit_dashboard_db';
const DB_VERSION = 3;

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('labels')) {
        db.createObjectStore('labels', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('categories')) {
        db.createObjectStore('categories', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('state')) {
        db.createObjectStore('state', { keyPath: 'key' });
      }
      if (!db.objectStoreNames.contains('pdfs')) {
        db.createObjectStore('pdfs', { keyPath: 'key' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function getAllFromStore(storeName) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const store = tx.objectStore(storeName);
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

async function replaceAllInStore(storeName, items) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    const store = tx.objectStore(storeName);
    store.clear();
    for (const item of items) {
      store.put(item);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function loadLabels() {
  try {
    return await getAllFromStore('labels');
  } catch (err) {
    console.error('Erro ao carregar etiquetas:', err);
    return [];
  }
}

export async function saveLabels(labels) {
  try {
    await replaceAllInStore('labels', labels);
  } catch (err) {
    console.error('Erro ao salvar etiquetas:', err);
  }
}

export async function loadCategories() {
  try {
    return await getAllFromStore('categories');
  } catch (err) {
    console.error('Erro ao carregar categorias:', err);
    return [];
  }
}

export async function saveCategories(categories) {
  try {
    await replaceAllInStore('categories', categories);
  } catch (err) {
    console.error('Erro ao salvar categorias:', err);
  }
}

export async function loadCart() {
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx = db.transaction('state', 'readonly');
      const store = tx.objectStore('state');
      const req = store.get('checkout_cart');
      req.onsuccess = () => resolve(req.result?.value || {});
      req.onerror = () => resolve({});
    });
  } catch {
    return {};
  }
}

export async function saveCart(cart) {
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx = db.transaction('state', 'readwrite');
      const store = tx.objectStore('state');
      store.put({ key: 'checkout_cart', value: cart });
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  } catch (err) {
    console.error('Erro ao salvar seleção:', err);
  }
}

export async function clearAllStorage() {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(['labels', 'state', 'pdfs'], 'readwrite');
      tx.objectStore('labels').clear();
      tx.objectStore('state').clear();
      tx.objectStore('pdfs').clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.error('Erro ao limpar banco:', err);
  }
}
