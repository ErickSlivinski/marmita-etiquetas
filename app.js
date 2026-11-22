/**
 * EtiquetaFit — Dashboard & Label Marketplace Controller
 * Standardized A4 Paisagem (297 × 210 mm) PDF export:
 * - Comum (6×3 cm): 4 cols × 6 rows = 24 labels/sheet
 * - Sucos (10×4 cm): 7 cols × 2 rows = 14 labels/sheet (rotated 90°, edge-to-edge)
 */

import {
  DEFAULT_CATEGORIES,
  LABEL_FORMATS,
  MAX_SHEETS_PER_LABEL,
  cleanOcrLabelName,
  detectFormatFromDimensions,
  resolveLabelFormat,
  getLabelFormatSpec,
  isSucosCategory,
  findSucosCategoryId,
  normalizeText,
  inferCategoryFromText,
  filterAndSortLabels,
  getCategoryCounts,
  clampSheets,
  calculateCheckoutSummary,
  createCategorySlug,
  validateBackupPayload
} from './lib/catalog.js?v=5';

import {
  loadLabels,
  saveLabels,
  loadCategories,
  saveCategories,
  loadCart,
  saveCart,
  clearAllStorage
} from './lib/storage.js?v=5';

import { generateDemoLabels } from './lib/demoLabels.js?v=5';

// Configure PDF.js worker if loaded
if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// ─── Application State ───────────────────────────────────────────────────────

let labels = [];
let categories = [];
let cart = {}; // { [labelId]: numberOfSheets }
let activeCategoryId = 'all';
let searchQuery = '';
let sortBy = 'name-asc';
let activeEditContext = null; // { mode: 'edit' | 'create', labelId?, dataUrl?, resolve? }
let activePreviewId = null;
let lastGeneratedPdfUrl = null;
let tesseractWorker = null;

// ─── DOM References ──────────────────────────────────────────────────────────

const sidebar = document.getElementById('sidebar');
const sidebarBackdrop = document.getElementById('sidebarBackdrop');
const menuBtn = document.getElementById('menuBtn');
const sidebarCloseBtn = document.getElementById('sidebarCloseBtn');

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const categoryNav = document.getElementById('categoryNav');
const quickFilterStrip = document.getElementById('quickFilterStrip');
const manageCategoriesBtn = document.getElementById('manageCategoriesBtn');

const activeViewTitle = document.getElementById('activeViewTitle');
const searchInput = document.getElementById('searchInput');
const clearSearchBtn = document.getElementById('clearSearchBtn');
const openCheckoutBtn = document.getElementById('openCheckoutBtn');
const checkoutCountBadge = document.getElementById('checkoutCountBadge');

const resultsSummary = document.getElementById('resultsSummary');
const selectionSummary = document.getElementById('selectionSummary');
const clearSelectionBtn = document.getElementById('clearSelectionBtn');
const sortSelect = document.getElementById('sortSelect');
const catalogGrid = document.getElementById('catalogGrid');
const emptyLibraryState = document.getElementById('emptyLibraryState');
const emptyFilterState = document.getElementById('emptyFilterState');
const emptyFilterText = document.getElementById('emptyFilterText');
const emptyUploadBtn = document.getElementById('emptyUploadBtn');
const emptyRestoreBackupBtn = document.getElementById('emptyRestoreBackupBtn');
const loadDemoBtn = document.getElementById('loadDemoBtn');
const resetFiltersBtn = document.getElementById('resetFiltersBtn');

// Checkout Drawer
const checkoutOverlay = document.getElementById('checkoutOverlay');
const closeCheckoutBtn = document.getElementById('closeCheckoutBtn');
const summaryModels = document.getElementById('summaryModels');
const summarySheets = document.getElementById('summarySheets');
const summaryStickers = document.getElementById('summaryStickers');
const checkoutItemsList = document.getElementById('checkoutItemsList');
const exportSuccessBox = document.getElementById('exportSuccessBox');
const exportSuccessMeta = document.getElementById('exportSuccessMeta');
const exportDownloadAgainLink = document.getElementById('exportDownloadAgainLink');
const exportFilenameInput = document.getElementById('exportFilenameInput');
const printSpecNote = document.getElementById('printSpecNote');
const drawerClearBtn = document.getElementById('drawerClearBtn');
const generatePdfBtn = document.getElementById('generatePdfBtn');

// Edit Label Modal
const editLabelModal = document.getElementById('editLabelModal');
const editModalTitle = document.getElementById('editModalTitle');
const editModalThumb = document.getElementById('editModalThumb');
const editLabelForm = document.getElementById('editLabelForm');
const editNameInput = document.getElementById('editNameInput');
const editCategorySelect = document.getElementById('editCategorySelect');
const editFormatSelect = document.getElementById('editFormatSelect');
const editFormatHint = document.getElementById('editFormatHint');
const quickNewCatInput = document.getElementById('quickNewCatInput');
const quickAddCatBtn = document.getElementById('quickAddCatBtn');
const closeEditModal = document.getElementById('closeEditModal');
const cancelEditBtn = document.getElementById('cancelEditBtn');

// Categories Manager Modal
const categoriesModal = document.getElementById('categoriesModal');
const closeCategoriesModal = document.getElementById('closeCategoriesModal');
const addCategoryForm = document.getElementById('addCategoryForm');
const newCategoryInput = document.getElementById('newCategoryInput');
const categoriesManagerList = document.getElementById('categoriesManagerList');

// Preview Modal
const previewModal = document.getElementById('previewModal');
const previewImg = document.getElementById('previewImg');
const previewName = document.getElementById('previewName');
const previewMeta = document.getElementById('previewMeta');
const closePreview = document.getElementById('closePreview');
const previewEditBtn = document.getElementById('previewEditBtn');
const previewAddSheetBtn = document.getElementById('previewAddSheetBtn');

// Settings / Backup Modal
const settingsBtn = document.getElementById('settingsBtn');
const mobileBackupBtn = document.getElementById('mobileBackupBtn');
const settingsModal = document.getElementById('settingsModal');
const closeSettings = document.getElementById('closeSettings');
const btnRestoreBundledBackup = document.getElementById('btnRestoreBundledBackup');
const btnExportBackup = document.getElementById('btnExportBackup');
const btnImportBackup = document.getElementById('btnImportBackup');
const importInput = document.getElementById('importInput');
const btnDangerClearAll = document.getElementById('btnDangerClearAll');
const clearConfirmBox = document.getElementById('clearConfirmBox');
const clearConfirmInput = document.getElementById('clearConfirmInput');
const cancelClearBtn = document.getElementById('cancelClearBtn');
const confirmClearBtn = document.getElementById('confirmClearBtn');

// Progress & Toasts
const progressOverlay = document.getElementById('progressOverlay');
const progressText = document.getElementById('progressText');
const progressBar = document.getElementById('progressBar');
const toastRegion = document.getElementById('toastRegion');

// ─── Initialization ──────────────────────────────────────────────────────────

async function init() {
  const [loadedLabels, loadedCategories, loadedCart] = await Promise.all([
    loadLabels(),
    loadCategories(),
    loadCart()
  ]);

  const rawCategories = loadedCategories.length > 0 ? loadedCategories : [...DEFAULT_CATEGORIES];
  categories = rawCategories.map(c => ({
    ...c,
    defaultFormat: LABEL_FORMATS[c.defaultFormat]
      ? c.defaultFormat
      : (/\bsucos?\b/.test(normalizeText(c.name)) || c.id === 'sucos' ? '10x4' : '6x3')
  }));
  if (loadedCategories.length === 0) {
    await saveCategories(categories);
  }

  labels = loadedLabels.map(l => ({
    ...l,
    categoryId: l.categoryId || 'sem-categoria',
    format: LABEL_FORMATS[l.format] ? l.format : undefined,
    createdAt: l.createdAt || Date.now()
  }));

  // Clean cart entries for deleted labels
  cart = {};
  for (const [id, qty] of Object.entries(loadedCart || {})) {
    if (labels.some(l => l.id === id) && clampSheets(qty) > 0) {
      cart[id] = clampSheets(qty);
    }
  }

  // Optional demo auto-seed via ?demo=1 when empty
  const params = new URLSearchParams(window.location.search);
  if (params.get('demo') === '1' && labels.length === 0) {
    labels = generateDemoLabels();
    await saveLabels(labels);
  }

  renderAll();
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function getCategoryName(categoryId) {
  const found = categories.find(c => c.id === categoryId);
  return found ? found.name : 'Sem Categoria';
}

function showToast(message, isError = false) {
  const toast = document.createElement('div');
  toast.className = `toast-item${isError ? ' error' : ''}`;
  toast.textContent = message;
  toastRegion.appendChild(toast);
  setTimeout(() => {
    toast.remove();
  }, 3800);
}

function showProgress(text, pct = 0) {
  progressText.textContent = text;
  progressBar.style.width = `${Math.min(100, Math.max(0, pct))}%`;
  progressOverlay.hidden = false;
}

function updateProgress(pct, text) {
  if (text) progressText.textContent = text;
  progressBar.style.width = `${Math.min(100, Math.max(0, pct))}%`;
}

function hideProgress() {
  progressOverlay.hidden = true;
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ─── Responsive Sidebar ──────────────────────────────────────────────────────

function openMobileSidebar() {
  sidebar.classList.add('open');
  sidebarBackdrop.hidden = false;
}

function closeMobileSidebar() {
  sidebar.classList.remove('open');
  sidebarBackdrop.hidden = true;
}

menuBtn.addEventListener('click', openMobileSidebar);
sidebarCloseBtn.addEventListener('click', closeMobileSidebar);
sidebarBackdrop.addEventListener('click', closeMobileSidebar);

// ─── Rendering ───────────────────────────────────────────────────────────────

function renderAll() {
  renderCategoriesNav();
  renderQuickFilterStrip();
  renderCatalogGrid();
  renderCheckoutSummary();
}

function renderCategoriesNav() {
  const counts = getCategoryCounts(labels, categories);
  categoryNav.innerHTML = '';

  const allBtn = document.createElement('button');
  allBtn.type = 'button';
  allBtn.className = `cat-nav-btn${activeCategoryId === 'all' ? ' active' : ''}`;
  allBtn.innerHTML = `
    <span class="cat-nav-name">Todas as Etiquetas</span>
    <span class="cat-nav-count tabular-nums">${counts.all || 0}</span>
  `;
  allBtn.addEventListener('click', () => {
    activeCategoryId = 'all';
    closeMobileSidebar();
    renderAll();
  });
  categoryNav.appendChild(allBtn);

  for (const cat of categories) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `cat-nav-btn${activeCategoryId === cat.id ? ' active' : ''}`;
    btn.innerHTML = `
      <span class="cat-nav-name">${escHtml(cat.name)}</span>
      <span class="cat-nav-count tabular-nums">${counts[cat.id] || 0}</span>
    `;
    btn.addEventListener('click', () => {
      activeCategoryId = cat.id;
      closeMobileSidebar();
      renderAll();
    });
    categoryNav.appendChild(btn);
  }

  activeViewTitle.textContent =
    activeCategoryId === 'all'
      ? 'Catálogo de Etiquetas'
      : getCategoryName(activeCategoryId);
}

function renderQuickFilterStrip() {
  const counts = getCategoryCounts(labels, categories);
  quickFilterStrip.innerHTML = '';

  const tabs = [{ id: 'all', name: 'Todas' }, ...categories];
  for (const tab of tabs) {
    const count = tab.id === 'all' ? counts.all : (counts[tab.id] || 0);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.role = 'tab';
    btn.ariaSelected = activeCategoryId === tab.id ? 'true' : 'false';
    btn.className = `filter-tab-btn${activeCategoryId === tab.id ? ' active' : ''}`;
    btn.textContent = `${tab.name} (${count})`;
    btn.addEventListener('click', () => {
      activeCategoryId = tab.id;
      renderAll();
    });
    quickFilterStrip.appendChild(btn);
  }
}

function renderCatalogGrid() {
  const enrichedLabels = labels.map(l => ({
    ...l,
    categoryName: getCategoryName(l.categoryId)
  }));

  const visibleLabels = filterAndSortLabels(enrichedLabels, {
    query: searchQuery,
    categoryId: activeCategoryId,
    sortBy,
    cart
  });

  resultsSummary.textContent =
    visibleLabels.length === labels.length
      ? `${labels.length} ${labels.length === 1 ? 'etiqueta' : 'etiquetas'}`
      : `${visibleLabels.length} de ${labels.length} etiquetas`;

  catalogGrid.innerHTML = '';

  if (labels.length === 0) {
    catalogGrid.hidden = true;
    emptyFilterState.hidden = true;
    emptyLibraryState.hidden = false;
    return;
  }

  emptyLibraryState.hidden = true;

  if (visibleLabels.length === 0) {
    catalogGrid.hidden = true;
    emptyFilterState.hidden = false;
    emptyFilterText.textContent = searchQuery
      ? `Nenhuma etiqueta encontrada para "${searchQuery}".`
      : 'Nenhuma etiqueta cadastrada nesta categoria.';
    return;
  }

  emptyFilterState.hidden = true;
  catalogGrid.hidden = false;

  for (const label of visibleLabels) {
    const qty = clampSheets(cart[label.id]);
    const card = document.createElement('article');
    card.className = `label-card${qty > 0 ? ' selected' : ''}`;
    card.dataset.id = label.id;

    const spec = getLabelFormatSpec(label, categories);
    const totalUnits = qty * spec.labelsPerSheet;
    const catName = getCategoryName(label.categoryId);

    card.innerHTML = `
      <button type="button" class="card-preview-btn" data-action="preview" data-id="${label.id}" aria-label="Visualizar ${escHtml(label.name)}">
        <img src="${label.dataUrl}" alt="${escHtml(label.name)}" />
      </button>
      <div class="card-body">
        <div class="card-meta-row">
          <span class="card-meta-text">${escHtml(catName)} · ${spec.shortDim}</span>
          <div class="card-icon-actions">
            <button type="button" class="icon-btn" data-action="edit" data-id="${label.id}" aria-label="Editar ${escHtml(label.name)}" title="Editar nome, categoria e tamanho">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
            </button>
            <button type="button" class="icon-btn danger" data-action="delete" data-id="${label.id}" aria-label="Excluir ${escHtml(label.name)}" title="Excluir etiqueta">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
            </button>
          </div>
        </div>
        <h3 class="card-title" title="${escHtml(label.name)}">${escHtml(label.name)}</h3>
        <div class="card-purchase-row">
          ${
            qty === 0
              ? `<button type="button" class="btn-add-label" data-action="add" data-id="${label.id}">Adicionar à folha</button>`
              : `<div class="qty-stepper">
                   <button type="button" class="qty-btn" data-action="dec" data-id="${label.id}" aria-label="Diminuir folhas de ${escHtml(label.name)}">−</button>
                   <input type="number" class="qty-input tabular-nums" min="0" max="${MAX_SHEETS_PER_LABEL}" value="${qty}" data-action="qty-input" data-id="${label.id}" aria-label="Quantidade de folhas para ${escHtml(label.name)}" />
                   <button type="button" class="qty-btn" data-action="inc" data-id="${label.id}" aria-label="Aumentar folhas de ${escHtml(label.name)}">+</button>
                 </div>
                 <span class="qty-summary tabular-nums">${qty} ${qty === 1 ? 'folha' : 'folhas'} · ${totalUnits} un.</span>`
          }
        </div>
      </div>
    `;

    catalogGrid.appendChild(card);
  }
}

function renderCheckoutSummary() {
  const summary = calculateCheckoutSummary(labels, cart, categories);
  checkoutCountBadge.textContent = `${summary.totalSheets} ${summary.totalSheets === 1 ? 'folha' : 'folhas'}`;

  if (summary.totalSheets === 0) {
    selectionSummary.textContent = 'Nenhuma folha selecionada';
    clearSelectionBtn.hidden = true;
  } else {
    selectionSummary.textContent = `${summary.totalModels} ${summary.totalModels === 1 ? 'modelo' : 'modelos'} · ${summary.totalSheets} ${summary.totalSheets === 1 ? 'folha' : 'folhas'} (${summary.totalStickers} etiquetas)`;
    clearSelectionBtn.hidden = false;
  }

  summaryModels.textContent = String(summary.totalModels);
  summarySheets.textContent = String(summary.totalSheets);
  summaryStickers.textContent = String(summary.totalStickers);

  const hasStd = summary.sheetsByFormat['6x3'] > 0;
  const hasJuice = summary.sheetsByFormat['10x4'] > 0;
  if (hasStd && hasJuice) {
    printSpecNote.textContent =
      'A4 Paisagem (297 × 210 mm): Comuns 4×6 (24/folha · 6×3 cm) + Sucos 7×2 (14/folha · 10×4 cm)';
  } else if (hasJuice) {
    printSpecNote.textContent =
      'Formato A4 Paisagem (297 × 210 mm) · 7 × 2 = 14 etiquetas/folha (10×4 cm) coladas com guias de corte';
  } else {
    printSpecNote.textContent =
      'Formato A4 Paisagem (297 × 210 mm) · 4 × 6 = 24 etiquetas/folha (6×3 cm) com guias de corte';
  }

  generatePdfBtn.disabled = summary.totalSheets === 0;
  drawerClearBtn.disabled = summary.totalSheets === 0;

  if (!checkoutOverlay.hidden) {
    renderCheckoutItems(summary);
  }
}

function renderCheckoutItems(summary) {
  checkoutItemsList.innerHTML = '';

  if (summary.items.length === 0) {
    checkoutItemsList.innerHTML = `
      <div class="checkout-empty">
        <p>Nenhuma etiqueta selecionada no momento.</p>
        <p>Volte ao catálogo e escolha a quantidade de folhas desejada.</p>
      </div>
    `;
    return;
  }

  for (const item of summary.items) {
    const row = document.createElement('div');
    row.className = 'checkout-item';
    const catName = getCategoryName(item.label.categoryId);
    const spec = LABEL_FORMATS[item.format] || LABEL_FORMATS['6x3'];

    row.innerHTML = `
      <img class="checkout-thumb" src="${item.label.dataUrl}" alt="${escHtml(item.label.name)}" />
      <div class="checkout-item-info">
        <div class="checkout-item-name" title="${escHtml(item.label.name)}">${escHtml(item.label.name)}</div>
        <div class="checkout-item-meta tabular-nums">${escHtml(catName)} · ${spec.shortDim} · ${item.totalStickers} etiquetas</div>
      </div>
      <div class="checkout-item-controls">
        <div class="qty-stepper">
          <button type="button" class="qty-btn" data-checkout-action="dec" data-id="${item.label.id}" aria-label="Diminuir folhas de ${escHtml(item.label.name)}">−</button>
          <input type="number" class="qty-input tabular-nums" min="0" max="${MAX_SHEETS_PER_LABEL}" value="${item.sheets}" data-checkout-action="qty" data-id="${item.label.id}" aria-label="Folhas de ${escHtml(item.label.name)}" />
          <button type="button" class="qty-btn" data-checkout-action="inc" data-id="${item.label.id}" aria-label="Aumentar folhas de ${escHtml(item.label.name)}">+</button>
        </div>
        <button type="button" class="icon-btn danger" data-checkout-action="remove" data-id="${item.label.id}" aria-label="Remover ${escHtml(item.label.name)} do checkout" title="Remover">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
        </button>
      </div>
    `;
    checkoutItemsList.appendChild(row);
  }
}

// ─── Cart / Selection Actions ────────────────────────────────────────────────

async function setLabelSheets(labelId, rawQty) {
  const qty = clampSheets(rawQty);
  if (qty <= 0) {
    delete cart[labelId];
  } else {
    cart[labelId] = qty;
  }
  exportSuccessBox.hidden = true;
  await saveCart(cart);
  renderCatalogGrid();
  renderCheckoutSummary();
}

async function clearCartSelection() {
  cart = {};
  exportSuccessBox.hidden = true;
  await saveCart(cart);
  renderCatalogGrid();
  renderCheckoutSummary();
}

catalogGrid.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-action]');
  if (!btn) return;
  const { action, id } = btn.dataset;

  if (action === 'preview') openPreviewModal(id);
  if (action === 'edit') openEditModalForLabel(id);
  if (action === 'delete') deleteLabel(id);
  if (action === 'add') setLabelSheets(id, 1);
  if (action === 'inc') setLabelSheets(id, (cart[id] || 0) + 1);
  if (action === 'dec') setLabelSheets(id, (cart[id] || 0) - 1);
});

catalogGrid.addEventListener('change', (e) => {
  const input = e.target.closest('[data-action="qty-input"]');
  if (!input) return;
  setLabelSheets(input.dataset.id, input.value);
});

// Caixa de confirmação reutilizável (usa as classes de modal já existentes)
function confirmDialog({ title, message, confirmText = 'Confirmar', cancelText = 'Cancelar', danger = false }) {
  return new Promise((resolve) => {
    const previouslyFocused = document.activeElement;
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.innerHTML = `
      <div class="modal" role="alertdialog" aria-modal="true" aria-labelledby="confirmTitle" aria-describedby="confirmMsg">
        <div class="modal-header">
          <h3 id="confirmTitle">${escHtml(title)}</h3>
        </div>
        <p id="confirmMsg" style="color: var(--text-secondary); margin-bottom: 18px;">${escHtml(message)}</p>
        <div class="modal-actions">
          <button type="button" class="btn-secondary" data-confirm="no">${escHtml(cancelText)}</button>
          <button type="button" class="${danger ? 'btn-danger-solid' : 'btn-primary'}" data-confirm="yes">${escHtml(confirmText)}</button>
        </div>
      </div>`;

    const cancelBtn = overlay.querySelector('[data-confirm="no"]');
    const okBtn = overlay.querySelector('[data-confirm="yes"]');

    function close(result) {
      document.removeEventListener('keydown', onKey, true);
      overlay.remove();
      if (previouslyFocused && typeof previouslyFocused.focus === 'function') previouslyFocused.focus();
      resolve(result);
    }

    function onKey(e) {
      if (e.key === 'Escape') {
        // fecha só a confirmação, sem fechar o checkout que está por trás
        e.preventDefault();
        e.stopImmediatePropagation();
        close(false);
      } else if (e.key === 'Tab') {
        // mantém o foco dentro da caixa
        e.preventDefault();
        (document.activeElement === cancelBtn ? okBtn : cancelBtn).focus();
      }
    }

    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close(false);
      const btn = e.target.closest('[data-confirm]');
      if (btn) close(btn.dataset.confirm === 'yes');
    });

    document.addEventListener('keydown', onKey, true);
    document.body.appendChild(overlay);
    cancelBtn.focus(); // foco no "Cancelar": evita confirmar sem querer com Enter
  });
}

async function requestClearSelection() {
  const summary = calculateCheckoutSummary(labels, cart, categories);
  if (summary.totalSheets === 0) return;

  const sheets = `${summary.totalSheets} ${summary.totalSheets === 1 ? 'folha' : 'folhas'}`;
  const models = `${summary.totalModels} ${summary.totalModels === 1 ? 'modelo' : 'modelos'}`;

  const ok = await confirmDialog({
    title: 'Limpar seleção?',
    message: `Isso remove ${sheets} (${models}, ${summary.totalStickers} etiquetas) da seleção. As etiquetas continuam no catálogo.`,
    confirmText: 'Limpar seleção',
    cancelText: 'Cancelar',
    danger: true
  });
  if (!ok) return;

  await clearCartSelection();
  showToast('Seleção limpa.');
}

clearSelectionBtn.addEventListener('click', requestClearSelection);
drawerClearBtn.addEventListener('click', requestClearSelection);

checkoutItemsList.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-checkout-action]');
  if (!btn) return;
  const { checkoutAction, id } = btn.dataset;

  if (checkoutAction === 'inc') setLabelSheets(id, (cart[id] || 0) + 1);
  if (checkoutAction === 'dec') setLabelSheets(id, (cart[id] || 0) - 1);
  if (checkoutAction === 'remove') setLabelSheets(id, 0);
});

checkoutItemsList.addEventListener('change', (e) => {
  const input = e.target.closest('[data-checkout-action="qty"]');
  if (!input) return;
  setLabelSheets(input.dataset.id, input.value);
});

// ─── Search & Sort ───────────────────────────────────────────────────────────

searchInput.addEventListener('input', () => {
  searchQuery = searchInput.value.trim();
  clearSearchBtn.hidden = searchQuery.length === 0;
  renderCatalogGrid();
});

clearSearchBtn.addEventListener('click', () => {
  searchInput.value = '';
  searchQuery = '';
  clearSearchBtn.hidden = true;
  renderCatalogGrid();
  searchInput.focus();
});

sortSelect.addEventListener('change', () => {
  sortBy = sortSelect.value;
  renderCatalogGrid();
});

resetFiltersBtn.addEventListener('click', () => {
  searchInput.value = '';
  searchQuery = '';
  clearSearchBtn.hidden = true;
  activeCategoryId = 'all';
  renderAll();
});

document.addEventListener('keydown', (e) => {
  if (e.key === '/' && !['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName)) {
    e.preventDefault();
    searchInput.focus();
  }
  if (e.key === 'Escape') {
    if (!checkoutOverlay.hidden) closeCheckoutDrawer();
    if (!editLabelModal.hidden) closeEditLabelModal();
    if (!categoriesModal.hidden) categoriesModal.hidden = true;
    if (!previewModal.hidden) previewModal.hidden = true;
    if (!settingsModal.hidden) settingsModal.hidden = true;
  }
});

// ─── Empty State Demo & Upload ───────────────────────────────────────────────

emptyUploadBtn.addEventListener('click', () => fileInput.click());

loadDemoBtn.addEventListener('click', async () => {
  const samples = generateDemoLabels();
  labels = [...labels, ...samples];
  await saveLabels(labels);
  renderAll();
  showToast(`${samples.length} etiquetas de exemplo carregadas no catálogo.`);
});

// ─── File Upload (Images, PDF, ZIP) + OCR ────────────────────────────────────

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
  handleIncomingFiles(Array.from(e.dataTransfer.files));
});

fileInput.addEventListener('change', () => {
  handleIncomingFiles(Array.from(fileInput.files));
  fileInput.value = '';
});

async function handleIncomingFiles(files) {
  if (!files.length) return;

  const extractedItems = []; // Array<{ dataUrl, fallbackName }>

  try {
    for (const file of files) {
      const lower = file.name.toLowerCase();
      if (lower.endsWith('.zip')) {
        const fromZip = await extractItemsFromZip(file);
        extractedItems.push(...fromZip);
      } else if (lower.endsWith('.pdf') || file.type === 'application/pdf') {
        const fromPdf = await extractItemsFromPdf(file);
        extractedItems.push(...fromPdf);
      } else if (file.type.startsWith('image/')) {
        const dataUrl = await blobToDataUrl(file);
        const fallbackName = cleanFilename(file.name);
        extractedItems.push({ dataUrl, fallbackName });
      }
    }

    if (extractedItems.length === 0) {
      showToast('Nenhum arquivo PNG, JPG, WEBP, PDF ou ZIP válido encontrado.', true);
      return;
    }

    if (extractedItems.length === 1) {
      await processSingleItemWithConfirm(extractedItems[0]);
    } else {
      await processBatchItems(extractedItems);
    }
  } catch (err) {
    console.error('Erro no upload:', err);
    showToast('Não foi possível processar um dos arquivos enviados.', true);
  } finally {
    hideProgress();
    if (tesseractWorker) {
      await tesseractWorker.terminate().catch(() => {});
      tesseractWorker = null;
    }
  }
}

function cleanFilename(filename) {
  return filename
    .split('/')
    .pop()
    .replace(/\.[^.]+$/, '')
    .replace(/[-_]+/g, ' ')
    .trim() || 'Etiqueta';
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function extractItemsFromZip(zipFile) {
  showProgress(`Extraindo ${zipFile.name}...`, 10);
  const zip = await window.JSZip.loadAsync(zipFile);
  const items = [];
  const entries = [];

  zip.forEach((path, entry) => {
    if (entry.dir) return;
    const lower = path.toLowerCase();
    if (
      lower.endsWith('.png') ||
      lower.endsWith('.jpg') ||
      lower.endsWith('.jpeg') ||
      lower.endsWith('.webp') ||
      lower.endsWith('.pdf')
    ) {
      entries.push({ path, entry });
    }
  });

  for (let i = 0; i < entries.length; i++) {
    const { path, entry } = entries[i];
    updateProgress(10 + ((i + 1) / Math.max(1, entries.length)) * 30);
    const blob = await entry.async('blob');
    if (path.toLowerCase().endsWith('.pdf')) {
      const pdfItems = await extractItemsFromPdf(blob, cleanFilename(path));
      items.push(...pdfItems);
    } else {
      const dataUrl = await blobToDataUrl(blob);
      items.push({ dataUrl, fallbackName: cleanFilename(path) });
    }
  }

  return items;
}

async function extractItemsFromPdf(pdfBlobOrFile, baseNameOverride) {
  if (!window.pdfjsLib) {
    throw new Error('Leitor de PDF indisponível no momento.');
  }

  const baseName = baseNameOverride || cleanFilename(pdfBlobOrFile.name || 'etiqueta.pdf');
  showProgress(`Convertendo PDF: ${baseName}...`, 15);

  const buffer = await pdfBlobOrFile.arrayBuffer();
  const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
  const items = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: 2.0 });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');

    await page.render({ canvasContext: ctx, viewport }).promise;
    const dataUrl = canvas.toDataURL('image/png');
    const fallbackName = pdf.numPages > 1 ? `${baseName} pág ${pageNum}` : baseName;
    items.push({ dataUrl, fallbackName });
    updateProgress(15 + (pageNum / pdf.numPages) * 25);
  }

  return items;
}

async function getImageDimensions(dataUrl) {
  try {
    const img = await loadImage(dataUrl);
    return {
      width: img.naturalWidth || img.width || 0,
      height: img.naturalHeight || img.height || 0
    };
  } catch {
    return { width: 0, height: 0 };
  }
}

async function ensureSucosCategoryIfNeeded() {
  let sucosId = findSucosCategoryId(categories);
  if (sucosId) return sucosId;

  const newCat = {
    id: createCategorySlug('Sucos', categories),
    name: 'Sucos',
    defaultFormat: '10x4'
  };
  const semCatIdx = categories.findIndex(c => c.id === 'sem-categoria');
  if (semCatIdx >= 0) {
    categories.splice(semCatIdx, 0, newCat);
  } else {
    categories.push(newCat);
  }
  await saveCategories(categories);
  return newCat.id;
}

async function recognizeLabelMetadata(dataUrl, fallbackName) {
  const dims = await getImageDimensions(dataUrl);
  const detectedFormat = detectFormatFromDimensions(dims.width, dims.height);

  if (detectedFormat === '10x4') {
    await ensureSucosCategoryIfNeeded();
  }

  let ocrText = '';
  let name = fallbackName;
  try {
    if (!window.Tesseract) throw new Error('Tesseract não carregado');
    if (!tesseractWorker) {
      tesseractWorker = await window.Tesseract.createWorker('por');
    }
    const { data: { text } } = await tesseractWorker.recognize(dataUrl);
    ocrText = text || '';
    name = cleanOcrLabelName(ocrText, fallbackName);
  } catch (err) {
    console.warn('Fallback de OCR aplicado:', err);
  }

  const categoryId = inferCategoryFromText(`${name} ${ocrText}`, categories, detectedFormat);
  const finalFormat =
    detectedFormat === '10x4' || isSucosCategory(categoryId, categories)
      ? '10x4'
      : '6x3';

  return {
    name,
    categoryId,
    detectedFormat: finalFormat,
    autoDetectedBySize: detectedFormat === '10x4',
    dimensions: dims
  };
}

async function processSingleItemWithConfirm(item) {
  showProgress('Lendo conteúdo e dimensões da etiqueta...', 50);
  const { name, categoryId, detectedFormat, autoDetectedBySize, dimensions } =
    await recognizeLabelMetadata(item.dataUrl, item.fallbackName);
  hideProgress();

  await new Promise((resolve) => {
    openCreateConfirmModal({
      name,
      categoryId,
      format: detectedFormat,
      autoDetectedBySize,
      dimensions,
      dataUrl: item.dataUrl,
      onDone: resolve
    });
  });
}

async function processBatchItems(items) {
  let added = 0;
  let juiceCount = 0;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    showProgress(
      `Lendo etiqueta ${i + 1} de ${items.length}...`,
      ((i + 1) / items.length) * 100
    );

    const { name, categoryId, detectedFormat } = await recognizeLabelMetadata(item.dataUrl, item.fallbackName);
    const duplicate = labels.some(l => l.name.toLowerCase() === name.toLowerCase());

    if (!duplicate) {
      labels.unshift({
        id: crypto.randomUUID(),
        name,
        categoryId,
        format: detectedFormat,
        dataUrl: item.dataUrl,
        createdAt: Date.now() - i
      });
      added++;
      if (detectedFormat === '10x4') juiceCount++;
    }
  }

  await saveLabels(labels);
  renderAll();
  const juiceNote = juiceCount > 0 ? ` (${juiceCount} no formato Sucos 10×4 cm)` : '';
  showToast(`${added} ${added === 1 ? 'etiqueta importada' : 'etiquetas importadas'}${juiceNote}.`);
}

// ─── Edit / Create Confirm Label Modal ───────────────────────────────────────

function populateCategorySelect(selectedId) {
  editCategorySelect.innerHTML = '';
  for (const cat of categories) {
    const opt = document.createElement('option');
    opt.value = cat.id;
    opt.textContent = cat.name;
    if (cat.id === selectedId) opt.selected = true;
    editCategorySelect.appendChild(opt);
  }
}

function setFormatHint(text) {
  if (text) {
    editFormatHint.textContent = text;
    editFormatHint.hidden = false;
  } else {
    editFormatHint.textContent = '';
    editFormatHint.hidden = true;
  }
}

function openCreateConfirmModal({ name, categoryId, format, autoDetectedBySize, dimensions, dataUrl, onDone }) {
  const resolvedFormat = format || (isSucosCategory(categoryId, categories) ? '10x4' : '6x3');
  activeEditContext = { mode: 'create', dataUrl, onDone };
  editModalTitle.textContent = 'Confirmar Nova Etiqueta';
  editModalThumb.src = dataUrl;
  editNameInput.value = name;
  quickNewCatInput.value = '';
  populateCategorySelect(categoryId);
  editFormatSelect.value = resolvedFormat;

  if (autoDetectedBySize) {
    const ratioTxt = dimensions?.width && dimensions?.height
      ? ` (${dimensions.width}×${dimensions.height}px)`
      : '';
    setFormatHint(`Proporção 10×4 cm detectada automaticamente${ratioTxt} — categoria Sucos selecionada (14/folha · A4 Paisagem).`);
  } else if (resolvedFormat === '10x4') {
    setFormatHint('Formato Sucos 10×4 cm selecionado (14/folha · A4 Paisagem).');
  } else {
    setFormatHint('Tamanho padrão 6×3 cm detectado (24/folha · A4 Paisagem).');
  }

  editLabelModal.hidden = false;
  setTimeout(() => {
    editNameInput.focus();
    editNameInput.select();
  }, 40);
}

function openEditModalForLabel(id) {
  const label = labels.find(l => l.id === id);
  if (!label) return;
  activeEditContext = { mode: 'edit', labelId: id };
  editModalTitle.textContent = 'Editar Etiqueta';
  editModalThumb.src = label.dataUrl;
  editNameInput.value = label.name;
  quickNewCatInput.value = '';
  populateCategorySelect(label.categoryId || 'sem-categoria');
  editFormatSelect.value = resolveLabelFormat(label, categories);
  setFormatHint('');
  editLabelModal.hidden = false;
  setTimeout(() => {
    editNameInput.focus();
    editNameInput.select();
  }, 40);
}

function closeEditLabelModal() {
  editLabelModal.hidden = true;
  if (activeEditContext?.onDone) {
    activeEditContext.onDone();
  }
  activeEditContext = null;
}

closeEditModal.addEventListener('click', closeEditLabelModal);
cancelEditBtn.addEventListener('click', closeEditLabelModal);
editLabelModal.addEventListener('click', (e) => {
  if (e.target === editLabelModal) closeEditLabelModal();
});

editCategorySelect.addEventListener('change', () => {
  const selectedCatId = editCategorySelect.value;
  if (isSucosCategory(selectedCatId, categories)) {
    editFormatSelect.value = '10x4';
    setFormatHint('Categoria Sucos utiliza layout 10×4 cm (7×2 = 14/folha · A4 Paisagem).');
  } else {
    editFormatSelect.value = '6x3';
    setFormatHint('Categoria padrão utiliza layout 6×3 cm (4×6 = 24/folha · A4 Paisagem).');
  }
});

editFormatSelect.addEventListener('change', () => {
  if (editFormatSelect.value === '10x4') {
    setFormatHint('Layout 10×4 cm: 7×2 = 14 etiquetas por folha A4 Paisagem.');
  } else {
    setFormatHint('Layout 6×3 cm: 4×6 = 24 etiquetas por folha A4 Paisagem.');
  }
});

quickAddCatBtn.addEventListener('click', async () => {
  const rawName = quickNewCatInput.value.trim();
  if (!rawName) return;

  const newCat = await addCategory(rawName);
  if (newCat) {
    populateCategorySelect(newCat.id);
    if (newCat.defaultFormat === '10x4') {
      editFormatSelect.value = '10x4';
    }
    quickNewCatInput.value = '';
    showToast(`Categoria "${newCat.name}" criada.`);
  }
});

editLabelForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const cleanName = editNameInput.value.trim();
  if (!cleanName) return;

  const chosenCategoryId = editCategorySelect.value || 'sem-categoria';
  const chosenFormat = editFormatSelect.value === '10x4' ? '10x4' : '6x3';

  if (activeEditContext?.mode === 'create') {
    labels.unshift({
      id: crypto.randomUUID(),
      name: cleanName,
      categoryId: chosenCategoryId,
      format: chosenFormat,
      dataUrl: activeEditContext.dataUrl,
      createdAt: Date.now()
    });
    await saveLabels(labels);
    showToast(`Etiqueta "${cleanName}" adicionada ao catálogo.`);
  } else if (activeEditContext?.mode === 'edit') {
    const target = labels.find(l => l.id === activeEditContext.labelId);
    if (target) {
      target.name = cleanName;
      target.categoryId = chosenCategoryId;
      target.format = chosenFormat;
      await saveLabels(labels);
      showToast('Etiqueta atualizada.');
    }
  }

  closeEditLabelModal();
  renderAll();
});

async function deleteLabel(id) {
  const target = labels.find(l => l.id === id);
  labels = labels.filter(l => l.id !== id);
  delete cart[id];
  await Promise.all([saveLabels(labels), saveCart(cart)]);
  renderAll();
  if (target) {
    showToast(`"${target.name}" removida do catálogo.`);
  }
}

// ─── Category Management ─────────────────────────────────────────────────────

async function addCategory(rawName) {
  const name = rawName.trim();
  if (!name) return null;

  const exists = categories.some(c => c.name.toLowerCase() === name.toLowerCase());
  if (exists) {
    showToast('Já existe uma categoria com esse nome.', true);
    return null;
  }

  const newCat = {
    id: createCategorySlug(name, categories),
    name,
    defaultFormat: /\bsucos?\b/.test(normalizeText(name)) ? '10x4' : '6x3'
  };

  // Keep "Sem Categoria" at the end
  const semCatIdx = categories.findIndex(c => c.id === 'sem-categoria');
  if (semCatIdx >= 0) {
    categories.splice(semCatIdx, 0, newCat);
  } else {
    categories.push(newCat);
  }

  await saveCategories(categories);
  renderAll();
  return newCat;
}

function renderCategoriesManager() {
  const counts = getCategoryCounts(labels, categories);
  categoriesManagerList.innerHTML = '';

  for (const cat of categories) {
    const isProtected = cat.id === 'sem-categoria';
    const catFormat = LABEL_FORMATS[cat.defaultFormat]
      ? cat.defaultFormat
      : (isSucosCategory(cat.id, categories) ? '10x4' : '6x3');
    const row = document.createElement('div');
    row.className = 'cat-manager-row';

    row.innerHTML = `
      <input
        type="text"
        class="text-input flex-1"
        value="${escHtml(cat.name)}"
        data-cat-rename="${cat.id}"
        aria-label="Nome da categoria ${escHtml(cat.name)}"
        ${isProtected ? 'disabled' : ''}
      />
      <select
        class="text-input cat-format-select"
        data-cat-format="${cat.id}"
        aria-label="Layout padrão para ${escHtml(cat.name)}"
        title="Formato padrão de impressão"
      >
        <option value="6x3"${catFormat === '6x3' ? ' selected' : ''}>6×3 cm (24/fl)</option>
        <option value="10x4"${catFormat === '10x4' ? ' selected' : ''}>10×4 cm (14/fl)</option>
      </select>
      <span class="cat-manager-count tabular-nums">${counts[cat.id] || 0} itens</span>
      ${
        isProtected
          ? ''
          : `<button type="button" class="icon-btn danger" data-cat-delete="${cat.id}" aria-label="Excluir categoria ${escHtml(cat.name)}" title="Excluir categoria">
               <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
             </button>`
      }
    `;
    categoriesManagerList.appendChild(row);
  }
}

manageCategoriesBtn.addEventListener('click', () => {
  renderCategoriesManager();
  categoriesModal.hidden = false;
  setTimeout(() => newCategoryInput.focus(), 40);
});

closeCategoriesModal.addEventListener('click', () => {
  categoriesModal.hidden = true;
});

categoriesModal.addEventListener('click', (e) => {
  if (e.target === categoriesModal) categoriesModal.hidden = true;
});

addCategoryForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const name = newCategoryInput.value.trim();
  if (!name) return;
  const created = await addCategory(name);
  if (created) {
    newCategoryInput.value = '';
    renderCategoriesManager();
    showToast(`Categoria "${created.name}" adicionada.`);
  }
});

categoriesManagerList.addEventListener('change', async (e) => {
  const formatSelect = e.target.closest('[data-cat-format]');
  if (formatSelect) {
    const id = formatSelect.dataset.catFormat;
    const cat = categories.find(c => c.id === id);
    if (!cat) return;
    cat.defaultFormat = formatSelect.value === '10x4' ? '10x4' : '6x3';
    await saveCategories(categories);
    renderAll();
    showToast(`Layout da categoria "${cat.name}" atualizado.`);
    return;
  }

  const input = e.target.closest('[data-cat-rename]');
  if (!input) return;
  const id = input.dataset.catRename;
  const newName = input.value.trim();
  const cat = categories.find(c => c.id === id);
  if (!cat) return;

  if (!newName) {
    input.value = cat.name;
    return;
  }

  cat.name = newName;
  await saveCategories(categories);
  renderAll();
  showToast('Categoria renomeada.');
});

categoriesManagerList.addEventListener('click', async (e) => {
  const btn = e.target.closest('[data-cat-delete]');
  if (!btn) return;
  const id = btn.dataset.catDelete;
  const removed = categories.find(c => c.id === id);

  categories = categories.filter(c => c.id !== id);
  labels = labels.map(l => (l.categoryId === id ? { ...l, categoryId: 'sem-categoria' } : l));
  if (activeCategoryId === id) activeCategoryId = 'all';

  await Promise.all([saveCategories(categories), saveLabels(labels)]);
  renderCategoriesManager();
  renderAll();

  if (removed) {
    showToast(`Categoria "${removed.name}" removida. Etiquetas movidas para Sem Categoria.`);
  }
});

// ─── Preview Modal ───────────────────────────────────────────────────────────

function openPreviewModal(id) {
  const label = labels.find(l => l.id === id);
  if (!label) return;
  activePreviewId = id;
  const spec = getLabelFormatSpec(label, categories);
  previewImg.src = label.dataUrl;
  previewName.textContent = label.name;
  previewMeta.textContent = `${getCategoryName(label.categoryId)} · ${spec.shortDim} (${spec.labelsPerSheet}/folha)`;
  previewModal.hidden = false;
}

closePreview.addEventListener('click', () => {
  previewModal.hidden = true;
});

previewModal.addEventListener('click', (e) => {
  if (e.target === previewModal) previewModal.hidden = true;
});

previewEditBtn.addEventListener('click', () => {
  if (!activePreviewId) return;
  const id = activePreviewId;
  previewModal.hidden = true;
  openEditModalForLabel(id);
});

previewAddSheetBtn.addEventListener('click', async () => {
  if (!activePreviewId) return;
  await setLabelSheets(activePreviewId, (cart[activePreviewId] || 0) + 1);
  showToast('+1 folha adicionada ao checkout.');
});

// ─── Checkout Drawer & Combined PDF Generation ───────────────────────────────

function openCheckoutDrawer() {
  const today = new Date().toISOString().split('T')[0];
  if (!exportFilenameInput.value) {
    exportFilenameInput.value = `etiquetas_lote_${today}`;
  }
  const summary = calculateCheckoutSummary(labels, cart, categories);
  renderCheckoutItems(summary);
  checkoutOverlay.hidden = false;
}

function closeCheckoutDrawer() {
  checkoutOverlay.hidden = true;
}

openCheckoutBtn.addEventListener('click', openCheckoutDrawer);
closeCheckoutBtn.addEventListener('click', closeCheckoutDrawer);
checkoutOverlay.addEventListener('click', (e) => {
  if (e.target === checkoutOverlay) closeCheckoutDrawer();
});

generatePdfBtn.addEventListener('click', generateCombinedPdf);

async function generateCombinedPdf() {
  const summary = calculateCheckoutSummary(labels, cart, categories);
  if (summary.totalSheets === 0) return;

  if (!window.jspdf?.jsPDF) {
    showToast('Biblioteca jsPDF ainda não carregada.', true);
    return;
  }

  showProgress(`Gerando PDF com ${summary.totalSheets} folha(s)...`, 0);

  try {
    const { jsPDF } = window.jspdf;
    const firstSpec = LABEL_FORMATS[summary.items[0].format] || LABEL_FORMATS['6x3'];
    const doc = new jsPDF({ orientation: firstSpec.orientation, unit: 'mm', format: 'a4' });

    let currentSheet = 0;
    let isFirstPage = true;

    for (const item of summary.items) {
      const rawImg = await loadImage(item.label.dataUrl);
      const spec = LABEL_FORMATS[item.format] || LABEL_FORMATS['6x3'];
      const printableImage = spec.rotate90 ? rotateImage90(rawImg) : rawImg;

      for (let s = 0; s < item.sheets; s++) {
        if (!isFirstPage) {
          doc.addPage('a4', spec.orientation);
        }
        isFirstPage = false;

        for (let row = 0; row < spec.rows; row++) {
          for (let col = 0; col < spec.cols; col++) {
            const x = spec.marginX + col * (spec.labelW + spec.gapX);
            const y = spec.marginY + row * (spec.labelH + spec.gapY);
            doc.addImage(printableImage, 'PNG', x, y, spec.labelW, spec.labelH, undefined, 'FAST');
          }
        }

        // Light cutting guide lines
        doc.setDrawColor(200, 200, 200);
        doc.setLineWidth(0.1);
        if (spec.gapX === 0 && spec.gapY === 0) {
          for (let col = 0; col <= spec.cols; col++) {
            const x = spec.marginX + col * spec.labelW;
            doc.line(x, spec.marginY, x, spec.marginY + spec.rows * spec.labelH);
          }
          for (let row = 0; row <= spec.rows; row++) {
            const y = spec.marginY + row * spec.labelH;
            doc.line(spec.marginX, y, spec.marginX + spec.cols * spec.labelW, y);
          }
        } else {
          for (let row = 0; row < spec.rows; row++) {
            for (let col = 0; col < spec.cols; col++) {
              const x = spec.marginX + col * (spec.labelW + spec.gapX);
              const y = spec.marginY + row * (spec.labelH + spec.gapY);
              doc.rect(x, y, spec.labelW, spec.labelH);
            }
          }
        }

        currentSheet++;
        updateProgress((currentSheet / summary.totalSheets) * 100);
        await sleep(10);
      }
    }

    const rawName = (exportFilenameInput.value || 'etiquetas_impressao')
      .trim()
      .replace(/\.pdf$/i, '')
      .replace(/[^a-zA-Z0-9_-]/g, '_');
    const filename = `${rawName || 'etiquetas_impressao'}.pdf`;

    const pdfBlob = doc.output('blob');
    if (lastGeneratedPdfUrl) {
      URL.revokeObjectURL(lastGeneratedPdfUrl);
    }
    lastGeneratedPdfUrl = URL.createObjectURL(pdfBlob);

    // Trigger download
    const link = document.createElement('a');
    link.href = lastGeneratedPdfUrl;
    link.download = filename;
    link.click();

    exportDownloadAgainLink.href = lastGeneratedPdfUrl;
    exportDownloadAgainLink.download = filename;
    exportSuccessMeta.textContent = `${filename} · ${summary.totalSheets} folhas · ${summary.totalStickers} etiquetas`;
    exportSuccessBox.hidden = false;

    showToast(`PDF "${filename}" gerado com ${summary.totalSheets} folhas.`);
  } catch (err) {
    console.error('Erro ao gerar PDF:', err);
    showToast('Erro ao gerar o PDF completo. Tente novamente.', true);
  } finally {
    hideProgress();
  }
}

function rotateImage90(img) {
  const w = img.naturalWidth || img.width;
  const h = img.naturalHeight || img.height;
  const canvas = document.createElement('canvas');
  canvas.width = h;
  canvas.height = w;
  const ctx = canvas.getContext('2d');
  ctx.translate(h, 0);
  ctx.rotate(Math.PI / 2);
  ctx.drawImage(img, 0, 0, w, h);
  return canvas.toDataURL('image/png');
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}

// ─── Backup & Settings Modal ─────────────────────────────────────────────────

function openSettingsModal() {
  closeMobileSidebar();
  clearConfirmBox.hidden = true;
  clearConfirmInput.value = '';
  confirmClearBtn.disabled = true;
  settingsModal.hidden = false;
}

settingsBtn.addEventListener('click', openSettingsModal);
mobileBackupBtn?.addEventListener('click', openSettingsModal);

closeSettings.addEventListener('click', () => {
  settingsModal.hidden = true;
});

settingsModal.addEventListener('click', (e) => {
  if (e.target === settingsModal) settingsModal.hidden = true;
});

function downloadBackupFile() {
  const payload = {
    version: 2,
    exportedAt: new Date().toISOString(),
    categories,
    labels
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const date = new Date().toISOString().split('T')[0];

  const a = document.createElement('a');
  a.href = url;
  a.download = `etiquetafit_backup_${date}.json`;
  a.click();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

btnExportBackup.addEventListener('click', () => {
  if (labels.length === 0) {
    showToast('Não há etiquetas para exportar.', true);
    return;
  }
  downloadBackupFile();
  showToast('Backup JSON exportado.');
});

async function applyImportedBackup(parsed) {
  const { labels: importedLabels, categories: importedCategories } = validateBackupPayload(parsed);

  if (importedLabels.length === 0) {
    showToast('Nenhuma etiqueta válida encontrada no arquivo.', true);
    return;
  }

  for (const cat of importedCategories) {
    if (!categories.some(c => c.id === cat.id)) {
      const semCatIdx = categories.findIndex(c => c.id === 'sem-categoria');
      if (semCatIdx >= 0) {
        categories.splice(semCatIdx, 0, cat);
      } else {
        categories.push(cat);
      }
    }
  }

  let added = 0;
  for (const item of importedLabels) {
    const duplicate = labels.some(
      l => l.name.toLowerCase() === item.name.toLowerCase() && l.dataUrl.length === item.dataUrl.length
    );
    if (!duplicate) {
      labels.unshift({
        id: crypto.randomUUID(),
        name: item.name,
        categoryId: categories.some(c => c.id === item.categoryId) ? item.categoryId : 'sem-categoria',
        format: item.format,
        dataUrl: item.dataUrl,
        createdAt: item.createdAt || Date.now()
      });
      added++;
    }
  }

  await Promise.all([saveCategories(categories), saveLabels(labels)]);
  renderAll();
  settingsModal.hidden = true;
  showToast(
    added > 0
      ? `${added} ${added === 1 ? 'etiqueta importada' : 'etiquetas importadas'}.`
      : 'Todas as etiquetas do arquivo já estavam cadastradas.'
  );
}

async function restoreUpdatedBackupFromServer() {
  showProgress('Carregando backup atualizado (83 etiquetas)...', 40);
  try {
    const res = await fetch('public/etiquetafit_backup_atualizado.json?v=5', { cache: 'no-store' });
    if (!res.ok) throw new Error('Não foi possível baixar o backup atualizado.');
    const parsed = await res.json();
    updateProgress(85, 'Salvando etiquetas no catálogo...');
    await applyImportedBackup(parsed);
  } catch (err) {
    showToast(err.message || 'Erro ao importar backup atualizado.', true);
  } finally {
    hideProgress();
  }
}

btnRestoreBundledBackup.addEventListener('click', restoreUpdatedBackupFromServer);
emptyRestoreBackupBtn.addEventListener('click', restoreUpdatedBackupFromServer);

btnImportBackup.addEventListener('click', () => {
  importInput.click();
});

importInput.addEventListener('change', (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (ev) => {
    try {
      const parsed = JSON.parse(ev.target.result);
      await applyImportedBackup(parsed);
    } catch (err) {
      showToast(err.message || 'Erro ao importar arquivo JSON.', true);
    } finally {
      importInput.value = '';
    }
  };
  reader.readAsText(file);
});

btnDangerClearAll.addEventListener('click', () => {
  if (labels.length === 0) {
    showToast('A biblioteca já está vazia.', true);
    return;
  }
  clearConfirmBox.hidden = false;
  setTimeout(() => clearConfirmInput.focus(), 50);
});

clearConfirmInput.addEventListener('input', () => {
  confirmClearBtn.disabled = clearConfirmInput.value.trim().toUpperCase() !== 'APAGAR';
});

cancelClearBtn.addEventListener('click', () => {
  clearConfirmBox.hidden = true;
  clearConfirmInput.value = '';
});

confirmClearBtn.addEventListener('click', async () => {
  if (clearConfirmInput.value.trim().toUpperCase() !== 'APAGAR') return;

  downloadBackupFile();
  labels = [];
  cart = {};
  activeCategoryId = 'all';
  await clearAllStorage();
  renderAll();
  clearConfirmBox.hidden = true;
  settingsModal.hidden = true;
  showToast('Biblioteca limpa. Um backup de segurança foi baixado.');
});

// ─── Start ───────────────────────────────────────────────────────────────────

init();
