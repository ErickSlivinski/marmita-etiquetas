/**
 * EtiquetaFit — App Logic
 * Frozen meal label sheet generator
 * Layout: A4 landscape, 4 cols × 6 rows = 24 labels per sheet
 * Label size: 60mm × 30mm (6cm × 3cm)
 */

'use strict';

// ─── Constants ────────────────────────────────────────────────────────────────

// A4 landscape in mm
const A4_W_MM = 297;
const A4_H_MM = 210;

// Grid: 4 columns × 6 rows = 24 labels
const COLS = 4;
const ROWS = 6;
const LABELS_PER_SHEET = COLS * ROWS;

// Minimal safe margin for printers (5mm each side)
const MARGIN_MM = 5;

// Labels scaled to fill the page (minus margins)
const LABEL_W_MM = (A4_W_MM - 2 * MARGIN_MM) / COLS;  // ~71.75mm
const LABEL_H_MM = (A4_H_MM - 2 * MARGIN_MM) / ROWS;  // ~33.33mm
const MARGIN_X_MM = MARGIN_MM;
const MARGIN_Y_MM = MARGIN_MM;

// ─── State ────────────────────────────────────────────────────────────────────

/** @type {Array<{id: string, name: string, dataUrl: string}>} */
let labels = [];

let pendingRenameId = null;

// Removed localStorage persistence in favor of IndexedDB
const STORAGE_KEY = 'etiquetafit_labels';

// ─── PDF Cache (IndexedDB) ────────────────────────────────────────────────────

const PDF_DB_NAME = 'etiquetafit_pdf_cache';
const PDF_DB_VERSION = 2; // Incremented for labels store

function openPdfCacheDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(PDF_DB_NAME, PDF_DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('pdfs')) {
        db.createObjectStore('pdfs', { keyPath: 'key' });
      }
      if (!db.objectStoreNames.contains('labels')) {
        db.createObjectStore('labels', { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveLabelsToDB() {
  try {
    const db = await openPdfCacheDB();
    const tx = db.transaction('labels', 'readwrite');
    const store = tx.objectStore('labels');
    
    // Clear and re-add all (simple approach for sync)
    await new Promise((resolve, reject) => {
      const clearReq = store.clear();
      clearReq.onsuccess = resolve;
      clearReq.onerror = reject;
    });

    for (const label of labels) {
      store.add(label);
    }
    
    return new Promise((resolve) => {
      tx.oncomplete = resolve;
    });
  } catch (e) {
    console.error('Erro ao salvar etiquetas no DB:', e);
  }
}

async function loadLabelsFromDB() {
  try {
    const db = await openPdfCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('labels', 'readonly');
      const store = tx.objectStore('labels');
      const req = store.getAll();
      req.onsuccess = () => {
        labels = req.result || [];
        resolve();
      };
      req.onerror = () => reject(req.error);
    });
  } catch (e) {
    console.error('Erro ao carregar etiquetas do DB:', e);
    labels = [];
  }
}

async function cachePDF(labelId, sheets, blob, filename) {
  try {
    const db = await openPdfCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('pdfs', 'readwrite');
      const store = tx.objectStore('pdfs');
      store.put({ key: `${labelId}_${sheets}`, blob, filename, createdAt: Date.now() });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (e) {
    console.warn('Erro ao salvar PDF no cache:', e);
  }
}

async function getCachedPDF(labelId, sheets) {
  try {
    const db = await openPdfCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('pdfs', 'readonly');
      const store = tx.objectStore('pdfs');
      const req = store.get(`${labelId}_${sheets}`);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (e) {
    console.warn('Erro ao buscar PDF no cache:', e);
    return null;
  }
}

async function clearCachedPDFsForLabel(labelId) {
  try {
    const db = await openPdfCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('pdfs', 'readwrite');
      const store = tx.objectStore('pdfs');
      const req = store.openCursor();
      req.onsuccess = (e) => {
        const cursor = e.target.result;
        if (cursor) {
          if (cursor.key.startsWith(labelId + '_')) cursor.delete();
          cursor.continue();
        } else {
          resolve();
        }
      };
      req.onerror = () => reject(req.error);
    });
  } catch (e) {
    console.warn('Erro ao limpar cache:', e);
  }
}

async function clearAllCachedPDFs() {
  try {
    const db = await openPdfCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('pdfs', 'readwrite');
      const store = tx.objectStore('pdfs');
      store.clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (e) {
    console.warn('Erro ao limpar todo o cache:', e);
  }
}

// ─── DOM refs ─────────────────────────────────────────────────────────────────

const sidebar        = document.getElementById('sidebar');
const sidebarToggle  = document.getElementById('sidebarToggle');
const menuBtn        = document.getElementById('menuBtn');
const uploadArea     = document.getElementById('uploadArea');
const fileInput      = document.getElementById('fileInput');
const labelsGrid     = document.getElementById('labelsGrid');
const labelCount     = document.getElementById('labelCount');
const clearAll       = document.getElementById('clearAll');
const chatArea       = document.getElementById('chatArea');
const chatInput      = document.getElementById('chatInput');
const sendBtn        = document.getElementById('sendBtn');
const statusDot      = document.getElementById('statusDot');
const renameModal    = document.getElementById('renameModal');
const renameInput    = document.getElementById('renameInput');
const cancelRename   = document.getElementById('cancelRename');
const confirmRename  = document.getElementById('confirmRename');
const previewModal   = document.getElementById('previewModal');
const previewImg     = document.getElementById('previewImg');
const previewName    = document.getElementById('previewName');
const closePreview   = document.getElementById('closePreview');
const progressOverlay = document.getElementById('progressOverlay');
const progressText   = document.getElementById('progressText');
const progressBar    = document.getElementById('progressBar');

// ─── Sidebar toggle ───────────────────────────────────────────────────────────

function toggleSidebar() {
  const isMobile = window.innerWidth <= 768;
  if (isMobile) {
    sidebar.classList.toggle('open');
  } else {
    sidebar.classList.toggle('collapsed');
  }
}

menuBtn.addEventListener('click', toggleSidebar);
sidebarToggle.addEventListener('click', toggleSidebar);

// ─── File Upload ──────────────────────────────────────────────────────────────

uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
  uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('drag-over');
  processFiles(Array.from(e.dataTransfer.files));
});

fileInput.addEventListener('change', () => {
  processFiles(Array.from(fileInput.files));
  fileInput.value = '';
});

async function processFiles(files) {
  const zipFiles = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.zip'));
  const imageFiles = Array.from(files).filter(f => f.type.startsWith('image/'));

  // Process zip files first
  for (const zip of zipFiles) {
    await processZip(zip);
  }

  // Process image files
  if (imageFiles.length === 1) {
    // Single file: show rename modal
    await addLabelWithRename(imageFiles[0]);
  } else if (imageFiles.length > 1) {
    // Multiple files: auto-name from filename (batch mode)
    await addLabelsBatch(imageFiles);
  }

  if (!zipFiles.length && !imageFiles.length) {
    appendAssistantMessage('<p>⚠️ Nenhum arquivo de imagem ou .zip encontrado. Envie arquivos PNG, JPG, WEBP ou um .zip contendo imagens.</p>', 'error-bubble');
  }
}

async function processZip(zipFile) {
  showProgress(`Extraindo ${zipFile.name}...`, 0);

  try {
    const zip = await JSZip.loadAsync(zipFile);
    const imageEntries = [];

    zip.forEach((path, entry) => {
      if (entry.dir) return;
      const lower = path.toLowerCase();
      if (lower.endsWith('.png') || lower.endsWith('.jpg') || lower.endsWith('.jpeg') || lower.endsWith('.webp')) {
        imageEntries.push({ path, entry });
      }
    });

    if (!imageEntries.length) {
      hideProgress();
      appendAssistantMessage('<p>⚠️ O arquivo .zip não contém imagens (PNG, JPG, WEBP).</p>', 'error-bubble');
      return;
    }

    let added = 0;
    for (let i = 0; i < imageEntries.length; i++) {
      const { path, entry } = imageEntries[i];
      updateProgress(((i + 1) / imageEntries.length) * 100);

      const blob = await entry.async('blob');
      const dataUrl = await blobToDataUrl(blob);
      const id = crypto.randomUUID();

      // Extract name from file path inside zip
      const fileName = path.split('/').pop();
      const name = fileName.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ');

      // Check for duplicates by name
      if (!labels.some(l => l.name.toLowerCase() === name.toLowerCase())) {
        labels.push({ id, name, dataUrl });
        added++;
      }
    }

    saveLabelsToDB();
    renderLabelsGrid();
    hideProgress();

    appendAssistantMessage(`<p>📦 Arquivo <strong>${escHtml(zipFile.name)}</strong> extraído com sucesso!</p><p>✅ <strong>${added}</strong> etiqueta${added !== 1 ? 's' : ''} adicionada${added !== 1 ? 's' : ''}.</p><p class="hint-text">💡 Renomeie qualquer etiqueta clicando no ✏️ na sidebar.</p>`);
  } catch (err) {
    hideProgress();
    console.error('Erro ao processar zip:', err);
    appendAssistantMessage('<p>❌ Erro ao abrir o arquivo .zip. Verifique se o arquivo está válido.</p>', 'error-bubble');
  }
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// Single file upload — shows rename modal
function addLabelWithRename(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const dataUrl = e.target.result;
      const id = crypto.randomUUID();
      const defaultName = file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ');

      openRenameModal(id, defaultName, async () => {
        const label = { id, name: renameInput.value.trim() || defaultName, dataUrl };
        labels.push(label);
        await saveLabelsToDB();
        renderLabelsGrid();
        resolve();
      });
    };
    reader.readAsDataURL(file);
  });
}

// Batch upload — auto-names from filenames, no modal
async function addLabelsBatch(files) {
  showProgress(`Importando ${files.length} etiquetas...`, 0);
  let added = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    updateProgress(((i + 1) / files.length) * 100);

    const dataUrl = await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.readAsDataURL(file);
    });

    const id = crypto.randomUUID();
    const name = file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ');

    // Skip duplicates
    if (!labels.some(l => l.name.toLowerCase() === name.toLowerCase())) {
      labels.push({ id, name, dataUrl });
      added++;
    }
  }

  saveLabelsToDB();
  renderLabelsGrid();
  hideProgress();

  appendAssistantMessage(`<p>📤 Upload em lote concluído!</p><p>✅ <strong>${added}</strong> etiqueta${added !== 1 ? 's' : ''} adicionada${added !== 1 ? 's' : ''}.</p><p class="hint-text">💡 Renomeie qualquer etiqueta clicando no ✏️ na sidebar.</p>`);
}

// ─── Labels Grid ──────────────────────────────────────────────────────────────

function renderLabelsGrid() {
  labelCount.textContent = labels.length;

  if (!labels.length) {
    labelsGrid.innerHTML = '<div class="empty-state"><span>Nenhuma etiqueta ainda</span></div>';
    return;
  }

  labelsGrid.innerHTML = '';
  labels.forEach(label => {
    const card = document.createElement('div');
    card.className = 'label-card';
    card.dataset.id = label.id;
    card.innerHTML = `
      <img class="label-thumb" src="${label.dataUrl}" alt="${escHtml(label.name)}" />
      <div class="label-info">
        <div class="label-name" title="${escHtml(label.name)}">${escHtml(label.name)}</div>
        <div class="label-meta">6×3 cm</div>
      </div>
      <div class="label-actions">
        <button class="icon-btn" data-action="preview" data-id="${label.id}" title="Pré-visualizar">👁️</button>
        <button class="icon-btn" data-action="rename" data-id="${label.id}" title="Renomear">✏️</button>
        <button class="icon-btn danger" data-action="delete" data-id="${label.id}" title="Remover">🗑️</button>
      </div>
    `;
    labelsGrid.appendChild(card);
  });
}

labelsGrid.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-action]');
  if (!btn) return;
  const { action, id } = btn.dataset;

  if (action === 'preview') openPreview(id);
  if (action === 'rename') {
    const label = labels.find(l => l.id === id);
    if (label) openRenameModal(id, label.name, async () => {
      label.name = renameInput.value.trim() || label.name;
      await saveLabelsToDB();
      renderLabelsGrid();
    });
  }
  if (action === 'delete') deleteLabel(id);
});

async function deleteLabel(id) {
  labels = labels.filter(l => l.id !== id);
  await saveLabelsToDB();
  clearCachedPDFsForLabel(id); // limpa PDFs em cache dessa etiqueta
  renderLabelsGrid();
}

clearAll.addEventListener('click', async () => {
  if (labels.length === 0) return;
  if (confirm('Remover todas as etiquetas salvas?')) {
    labels = [];
    await saveLabelsToDB();
    clearAllCachedPDFs(); // limpa todos os PDFs em cache
    renderLabelsGrid();
  }
});

// ─── Preview modal ────────────────────────────────────────────────────────────

function openPreview(id) {
  const label = labels.find(l => l.id === id);
  if (!label) return;
  previewImg.src = label.dataUrl;
  previewName.textContent = label.name;
  previewModal.hidden = false;
}

closePreview.addEventListener('click', () => { previewModal.hidden = true; });
previewModal.addEventListener('click', (e) => { if (e.target === previewModal) previewModal.hidden = true; });

// ─── Rename modal ─────────────────────────────────────────────────────────────

let onRenameConfirm = null;

function openRenameModal(id, currentName, onConfirm) {
  pendingRenameId = id;
  renameInput.value = currentName;
  onRenameConfirm = onConfirm;
  renameModal.hidden = false;
  setTimeout(() => {
    renameInput.focus();
    renameInput.select();
  }, 50);
}

confirmRename.addEventListener('click', () => {
  renameModal.hidden = true;
  if (onRenameConfirm) { onRenameConfirm(); onRenameConfirm = null; }
});

cancelRename.addEventListener('click', () => {
  renameModal.hidden = true;
  onRenameConfirm = null;
});

renameInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') confirmRename.click();
  if (e.key === 'Escape') cancelRename.click();
});

// ─── Chat ─────────────────────────────────────────────────────────────────────

// Auto-resize textarea
chatInput.addEventListener('input', () => {
  chatInput.style.height = 'auto';
  chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
});

chatInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    handleSend();
  }
});

sendBtn.addEventListener('click', handleSend);

// Example chips
document.addEventListener('click', (e) => {
  if (e.target.matches('.example-chip')) {
    chatInput.value = e.target.dataset.example;
    chatInput.dispatchEvent(new Event('input'));
    chatInput.focus();
  }
});

async function handleSend() {
  const text = chatInput.value.trim();
  if (!text) return;

  chatInput.value = '';
  chatInput.style.height = 'auto';

  appendUserMessage(text);
  await processRequest(text);
}

function appendUserMessage(text) {
  const div = document.createElement('div');
  div.className = 'chat-message user';
  div.innerHTML = `
    <div class="avatar">🧑</div>
    <div class="bubble"><p>${escHtml(text)}</p></div>
  `;
  chatArea.appendChild(div);
  scrollChat();
}

function appendAssistantMessage(html, extraClass = '') {
  const div = document.createElement('div');
  div.className = `chat-message assistant${extraClass ? ' ' + extraClass : ''}`;
  div.innerHTML = `
    <div class="avatar">🤖</div>
    <div class="bubble">${html}</div>
  `;
  chatArea.appendChild(div);
  scrollChat();
  return div;
}

function scrollChat() {
  chatArea.scrollTo({ top: chatArea.scrollHeight, behavior: 'smooth' });
}

// ─── Natural language parsing ─────────────────────────────────────────────────

/**
 * Parse a request string to extract one or multiple label + quantity pairs.
 * Supports patterns like:
 *   "2 folhas de X"
 *   "quero 3 folhas do X"
 *   "preciso de 2 folhas da X e 1 folha do Y"
 *   "X - 3 folhas"
 * Returns Array<{labelName: string, sheets: number}>
 */
function parseRequest(text) {
  const lower = text.toLowerCase();
  const results = [];

  // Pattern 1: "N folha(s) de/do/da LABEL"
  // Pattern 2: "LABEL — N folha(s)"
  // We split by connectors (e e, mais, vírgula) to handle multiple requests

  // First, normalise separators
  const normalised = lower
    .replace(/\be\s+(\d)/g, '&&$1')   // "e 2 folhas" -> "&&2 folhas"
    .replace(/,\s*(\d)/g, '&&$1')     // ", 2 folhas" -> "&&2 folhas"
    .replace(/\s+mais\s+/g, '&&');    // "mais 2" -> "&&2"

  const parts = normalised.split('&&');

  for (const part of parts) {
    const match = extractLabelRequest(part.trim());
    if (match) results.push(match);
  }

  // Deduplicate by label
  const seen = new Set();
  return results.filter(r => {
    const key = r.labelName.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function extractLabelRequest(part) {
  // Patterns for quantity
  const qtyWords = { 'uma': 1, 'um': 1, 'dois': 2, 'duas': 2, 'tres': 3, 'três': 3,
    'quatro': 4, 'cinco': 5, 'seis': 6, 'sete': 7, 'oito': 8, 'nove': 9, 'dez': 10 };

  let sheets = null;
  let labelPart = part;

  // Try "N folha(s) ..." — number at start
  let m = part.match(/^(?:quero|preciso\s+de|me\s+d[eêá]|gera|gerar|fazer|fa[cç]a|imprimir)?\s*(\d+|[a-zêáãçé]+)\s+folha[s]?\s+(?:de|do|da|das|dos)?\s*(.*)/);
  if (m) {
    const rawQty = m[1];
    sheets = parseInt(rawQty) || qtyWords[rawQty] || 1;
    labelPart = m[2].trim();
  }

  // Try "... — N folha(s)" — number at end
  if (!sheets) {
    m = part.match(/(.+?)\s+[-–—]\s*(\d+)\s+folha[s]?/);
    if (m) {
      labelPart = m[1].trim();
      sheets = parseInt(m[2]);
    }
  }

  // Fallback: just a number somewhere
  if (!sheets) {
    m = part.match(/(\d+)\s+folha[s]?/);
    if (m) {
      sheets = parseInt(m[1]);
      labelPart = part.replace(m[0], '').trim();
    }
  }

  if (!sheets) return null;

  // Clean label name
  const cleaned = labelPart
    .replace(/^(de|do|da|das|dos|etiqueta|etiquetas|marmita|do\s+|da\s+)\s*/gi, '')
    .replace(/[.,;!?]+$/, '')
    .trim();

  if (!cleaned) return null;

  return { labelName: cleaned, sheets: Math.max(1, Math.min(sheets, 50)) };
}

/**
 * Find best matching label from library using fuzzy / keyword matching.
 * Returns the label object or null.
 */
function findLabel(queryName) {
  if (!labels.length) return null;

  const queryWords = normalise(queryName).split(/\s+/).filter(w => w.length > 2);

  let bestScore = 0;
  let bestLabel = null;

  for (const label of labels) {
    const labelNorm = normalise(label.name);
    const labelWords = labelNorm.split(/\s+/).filter(w => w.length > 2);

    // Exact substring match — highest priority
    if (labelNorm.includes(normalise(queryName))) {
      return label;
    }

    // Score based on common meaningful words
    let score = 0;
    for (const qw of queryWords) {
      if (labelNorm.includes(qw)) score += qw.length; // weight longer words more
    }
    // Bonus for reverse (label words in query)
    for (const lw of labelWords) {
      if (normalise(queryName).includes(lw)) score += lw.length * 0.5;
    }

    if (score > bestScore) {
      bestScore = score;
      bestLabel = label;
    }
  }

  // Require at least one meaningful word matched
  return bestScore >= 3 ? bestLabel : null;
}

function normalise(str) {
  return str
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove accents
    .replace(/[^a-z0-9\s]/g, '')
    .trim();
}

// ─── Request processing ───────────────────────────────────────────────────────

async function processRequest(text) {
  setBusy(true);

  // Show typing indicator
  const typingDiv = appendAssistantMessage(`
    <div class="typing-indicator">
      <div class="typing-dot"></div>
      <div class="typing-dot"></div>
      <div class="typing-dot"></div>
    </div>
  `);

  await sleep(700); // realistic delay

  typingDiv.remove();

  if (!labels.length) {
    appendAssistantMessage('<p>⚠️ Você ainda não fez upload de nenhuma etiqueta. Adicione suas etiquetas no painel à esquerda e depois me diga o que precisa.</p>', 'error-bubble');
    setBusy(false);
    return;
  }

  const requests = parseRequest(text);

  if (!requests.length) {
    appendAssistantMessage(`
      <p>🤔 Não consegui entender o pedido. Tente algo como:</p>
      <div class="examples">
        <button class="example-chip" data-example="Quero 2 folhas do escondidinho de mandioquinha">💬 "Quero 2 folhas do escondidinho de mandioquinha"</button>
        <button class="example-chip" data-example="3 folhas do filé de frango">💬 "3 folhas do filé de frango"</button>
      </div>
    `);
    setBusy(false);
    return;
  }

  // Resolve labels
  const resolved = [];
  const notFound = [];

  for (const req of requests) {
    const label = findLabel(req.labelName);
    if (label) {
      resolved.push({ label, sheets: req.sheets, queryName: req.labelName });
    } else {
      notFound.push(req.labelName);
    }
  }

  if (!resolved.length) {
    const names = notFound.map(n => `"${n}"`).join(', ');
    appendAssistantMessage(`<p>❌ Não encontrei nenhuma etiqueta correspondente a: ${escHtml(names)}.</p><p>Verifique os nomes das etiquetas salvas no painel ou tente usar palavras-chave do nome da etiqueta.</p>`, 'error-bubble');
    setBusy(false);
    return;
  }

  // Build confirmation message
  let confirmHtml = `<p>✅ Encontrei ${resolved.length === 1 ? 'a etiqueta' : 'as etiquetas'}! Gerando ${resolved.reduce((a, r) => a + r.sheets, 0)} folha(s) de PDF...</p>`;

  for (const r of resolved) {
    const total = r.sheets * LABELS_PER_SHEET;
    confirmHtml += `
      <div class="result-card">
        <img class="result-thumb" src="${r.label.dataUrl}" alt="${escHtml(r.label.name)}" />
        <div class="result-info">
          <div class="result-name">${escHtml(r.label.name)}</div>
          <div class="result-meta">${r.sheets} folha${r.sheets > 1 ? 's' : ''} · ${total} etiqueta${total > 1 ? 's' : ''} · A4 landscape</div>
        </div>
      </div>
    `;
  }

  if (notFound.length) {
    confirmHtml += `<p style="margin-top:12px; font-size:12px; color:var(--yellow)">⚠️ Não encontrei: ${notFound.map(n => `"${escHtml(n)}"`).join(', ')}</p>`;
  }

  const confirmDiv = appendAssistantMessage(confirmHtml);
  scrollChat();

  await sleep(400);

  // Generate PDFs
  try {
    for (const r of resolved) {
      await generateAndAttachPDF(r, confirmDiv);
    }
  } catch (err) {
    console.error('PDF generation error:', err);
    appendAssistantMessage('<p>❌ Ocorreu um erro ao gerar o PDF. Tente novamente.</p>', 'error-bubble');
  }

  setBusy(false);
}

// ─── PDF Generation ───────────────────────────────────────────────────────────

async function generateAndAttachPDF({ label, sheets }, parentDiv) {
  let pdfBlob, filename, fromCache = false;

  // 1. Check cache first
  const cached = await getCachedPDF(label.id, sheets);
  if (cached) {
    pdfBlob = cached.blob;
    filename = cached.filename;
    fromCache = true;
  }

  // 2. Generate new PDF if not cached
  if (!fromCache) {
    showProgress(`Gerando PDF: ${label.name}...`, 0);

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const img = await loadImage(label.dataUrl);

    for (let s = 0; s < sheets; s++) {
      if (s > 0) doc.addPage('a4', 'landscape');

      for (let row = 0; row < ROWS; row++) {
        for (let col = 0; col < COLS; col++) {
          const x = MARGIN_X_MM + col * LABEL_W_MM;
          const y = MARGIN_Y_MM + row * LABEL_H_MM;
          doc.addImage(img, 'PNG', x, y, LABEL_W_MM, LABEL_H_MM, undefined, 'FAST');
        }
      }

      // Draw faint grid lines to aid cutting
      doc.setDrawColor(200, 200, 200);
      doc.setLineWidth(0.1);

      for (let col = 0; col <= COLS; col++) {
        const x = MARGIN_X_MM + col * LABEL_W_MM;
        doc.line(x, MARGIN_Y_MM, x, MARGIN_Y_MM + ROWS * LABEL_H_MM);
      }
      for (let row = 0; row <= ROWS; row++) {
        const y = MARGIN_Y_MM + row * LABEL_H_MM;
        doc.line(MARGIN_X_MM, y, MARGIN_X_MM + COLS * LABEL_W_MM, y);
      }

      updateProgress(((s + 1) / sheets) * 100);
      await sleep(20);
    }

    hideProgress();

    const safeName = label.name.replace(/[^a-z0-9\s]/gi, '').replace(/\s+/g, '_').toLowerCase();
    filename = `etiqueta_${safeName}_${sheets}folha${sheets > 1 ? 's' : ''}.pdf`;
    pdfBlob = doc.output('blob');

    // 3. Save to cache for future use
    await cachePDF(label.id, sheets, pdfBlob, filename);
  }

  // 4. Attach download to chat UI
  const pdfUrl = URL.createObjectURL(pdfBlob);
  const bubble = parentDiv.querySelector('.bubble');
  const resultCard = bubble.querySelector(`.result-card img[src="${label.dataUrl}"]`)?.closest('.result-card');

  const btn = document.createElement('button');
  btn.className = 'download-btn';
  btn.innerHTML = fromCache ? `⚡ Baixar PDF (instantâneo)` : `⬇️ Baixar PDF`;
  btn.onclick = () => {
    const a = document.createElement('a');
    a.href = pdfUrl;
    a.download = filename;
    a.click();
  };

  if (resultCard) {
    resultCard.appendChild(btn);
  }

  // Success tag with cache indicator
  const successTag = document.createElement('div');
  successTag.className = 'success-tag';
  successTag.innerHTML = fromCache
    ? `⚡ ${filename} <span style="opacity:0.6; font-weight:400">(do cache — entrega instantânea)</span>`
    : `✅ ${filename}`;
  if (resultCard) resultCard.after(successTag);

  scrollChat();
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}

// ─── Progress ─────────────────────────────────────────────────────────────────

function showProgress(text, pct) {
  progressText.textContent = text;
  progressBar.style.width = pct + '%';
  progressOverlay.hidden = false;
}

function updateProgress(pct) {
  progressBar.style.width = pct + '%';
}

function hideProgress() {
  progressOverlay.hidden = true;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function setBusy(busy) {
  sendBtn.disabled = busy;
  chatInput.disabled = busy;
  statusDot.classList.toggle('busy', busy);
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function escHtml(str) {
  if (!str) return '';
  return str.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function appendAssistantMessage(html, className = 'assistant') {
  const div = document.createElement('div');
  div.className = `chat-message ${className}`;
  div.innerHTML = `
    <div class="avatar">${className === 'error-bubble' ? '⚠️' : '🤖'}</div>
    <div class="bubble">${html}</div>
  `;
  chatArea.appendChild(div);
  scrollChat();
  return div;
}

function scrollChat() {
  chatArea.scrollTop = chatArea.scrollHeight;
}

// ─── Init ─────────────────────────────────────────────────────────────────────

loadLabelsFromDB().then(() => {
  renderLabelsGrid();
});
chatInput.focus();
