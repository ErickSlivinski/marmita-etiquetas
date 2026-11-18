/**
 * Pure domain logic for EtiquetaFit Dashboard
 * Shared between browser UI (app.js) and Node unit tests (lib/catalog.test.js).
 */

export const LABELS_PER_SHEET = 24;
export const MAX_SHEETS_PER_LABEL = 50;

export const DEFAULT_CATEGORIES = [
  { id: 'low-carb', name: 'Low Carb' },
  { id: 'vegetariana', name: 'Vegetariana' },
  { id: 'frango', name: 'Frango & Aves' },
  { id: 'carnes', name: 'Carnes' },
  { id: 'massas', name: 'Massas & Escondidinhos' },
  { id: 'sem-categoria', name: 'Sem Categoria' }
];

export function normalizeText(str) {
  if (!str) return '';
  return String(str)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function cleanOcrLabelName(rawText, fallbackName = 'Etiqueta sem nome') {
  if (!rawText || typeof rawText !== 'string') return fallbackName;

  const ignoredPatterns = [
    /^peso\s*l[ií]q/i,
    /^val[:\s]/i,
    /^validade/i,
    /^fab[:\s]/i,
    /^lote/i,
    /^ingredientes/i,
    /^informa[cç][aã]o\s+nutricional/i,
    /^modo\s+de\s+preparo/i,
    /^conservar/i,
    /^\d+\s*(g|kg|ml|l|kcal)$/i
  ];

  const lines = rawText
    .split('\n')
    .map(l => l.trim().replace(/\s+/g, ' '))
    .filter(l => l.length > 2)
    .filter(l => !ignoredPatterns.some(rx => rx.test(l)));

  if (lines.length === 0) return fallbackName;

  let name = lines.slice(0, 2).join(' - ');
  name = name
    .toLowerCase()
    .replace(/(^|\s|-)([a-zà-öø-ÿ])/g, (_, sep, ch) => sep + ch.toUpperCase());

  return name.slice(0, 80).trim() || fallbackName;
}

export function inferCategoryFromText(text, categories = DEFAULT_CATEGORIES) {
  const norm = normalizeText(text);
  if (!norm) return 'sem-categoria';

  const findCatId = (targetSlug, keywordList) => {
    const byId = categories.find(c => c.id === targetSlug);
    if (byId) return byId.id;
    const byName = categories.find(c =>
      keywordList.some(k => normalizeText(c.name).includes(k))
    );
    return byName ? byName.id : null;
  };

  if (/\b(low\s*carb|cetogen|keto|fit|abobrinha|couve flor|couve-flor|berinjela)\b/.test(norm)) {
    const id = findCatId('low-carb', ['low carb', 'fit']);
    if (id) return id;
  }

  if (/\b(vegetarian|vegan|veggie|soja|grao de bico|lentilha|tofu|cogumelo|legumes|abobora|brocolis)\b/.test(norm) &&
      !/\b(frango|carne|bife|patinho|tilapia|salmao|peixe|porco|bacon)\b/.test(norm)) {
    const id = findCatId('vegetariana', ['vegetariana', 'vegana']);
    if (id) return id;
  }

  if (/\b(escondidinho|lasanha|nhoque|panqueca|macarrao|rondelli|canelone|pure|mandioquinha)\b/.test(norm)) {
    const id = findCatId('massas', ['massas', 'escondidinho']);
    if (id) return id;
  }

  if (/\b(frango|galinha|sobrecoxa|peito de frango|strogonoff de frango|parmegiana de frango)\b/.test(norm)) {
    const id = findCatId('frango', ['frango', 'aves']);
    if (id) return id;
  }

  if (/\b(carne|patinho|alcatra|file mignon|bife|almondega|kibe|moida|costela|ragu|picadinho|cupim)\b/.test(norm)) {
    const id = findCatId('carnes', ['carne', 'bovino']);
    if (id) return id;
  }

  return categories.some(c => c.id === 'sem-categoria')
    ? 'sem-categoria'
    : (categories[0]?.id ?? 'sem-categoria');
}

export function filterAndSortLabels(labels, options = {}) {
  const {
    query = '',
    categoryId = 'all',
    sortBy = 'name-asc',
    cart = {}
  } = options;

  const normQuery = normalizeText(query);
  const queryTerms = normQuery ? normQuery.split(' ').filter(Boolean) : [];

  const filtered = labels.filter(label => {
    if (categoryId !== 'all' && (label.categoryId || 'sem-categoria') !== categoryId) {
      return false;
    }

    if (queryTerms.length > 0) {
      const haystack = normalizeText(`${label.name} ${label.categoryName || ''}`);
      return queryTerms.every(term => haystack.includes(term));
    }

    return true;
  });

  const sorted = [...filtered];
  sorted.sort((a, b) => {
    if (sortBy === 'selected-first') {
      const qa = Number(cart[a.id] || 0);
      const qb = Number(cart[b.id] || 0);
      if (qa > 0 && qb === 0) return -1;
      if (qb > 0 && qa === 0) return 1;
      if (qa !== qb) return qb - qa;
    }
    if (sortBy === 'newest') {
      return (b.createdAt || 0) - (a.createdAt || 0);
    }
    if (sortBy === 'name-desc') {
      return b.name.localeCompare(a.name, 'pt-BR', { sensitivity: 'base' });
    }
    return a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'base' });
  });

  return sorted;
}

export function getCategoryCounts(labels, categories) {
  const counts = { all: labels.length };
  for (const cat of categories) {
    counts[cat.id] = 0;
  }
  for (const label of labels) {
    const cid = label.categoryId || 'sem-categoria';
    counts[cid] = (counts[cid] || 0) + 1;
  }
  return counts;
}

export function clampSheets(value) {
  const n = parseInt(value, 10);
  if (Number.isNaN(n) || n <= 0) return 0;
  return Math.min(n, MAX_SHEETS_PER_LABEL);
}

export function calculateCheckoutSummary(labels, cart, labelsPerSheet = LABELS_PER_SHEET) {
  const items = [];
  let totalSheets = 0;

  for (const label of labels) {
    const sheets = clampSheets(cart[label.id]);
    if (sheets > 0) {
      const totalStickers = sheets * labelsPerSheet;
      items.push({
        label,
        sheets,
        totalStickers
      });
      totalSheets += sheets;
    }
  }

  return {
    items,
    totalModels: items.length,
    totalSheets,
    totalStickers: totalSheets * labelsPerSheet
  };
}

export function createCategorySlug(name, existingCategories = []) {
  const base = normalizeText(name).replace(/\s+/g, '-') || 'categoria';
  let slug = base;
  let suffix = 2;
  while (existingCategories.some(c => c.id === slug)) {
    slug = `${base}-${suffix++}`;
  }
  return slug;
}

export function validateBackupPayload(payload) {
  if (Array.isArray(payload)) {
    // Legacy v1 array of labels
    const validLabels = payload.filter(item => item && item.name && item.dataUrl);
    return {
      labels: validLabels.map(l => ({
        name: String(l.name),
        dataUrl: String(l.dataUrl),
        categoryId: l.categoryId || 'sem-categoria',
        createdAt: l.createdAt || Date.now()
      })),
      categories: []
    };
  }

  if (payload && typeof payload === 'object' && Array.isArray(payload.labels)) {
    const validLabels = payload.labels.filter(item => item && item.name && item.dataUrl);
    const validCategories = Array.isArray(payload.categories)
      ? payload.categories.filter(c => c && c.id && c.name)
      : [];

    return {
      labels: validLabels.map(l => ({
        name: String(l.name),
        dataUrl: String(l.dataUrl),
        categoryId: l.categoryId || 'sem-categoria',
        createdAt: l.createdAt || Date.now()
      })),
      categories: validCategories.map(c => ({
        id: String(c.id),
        name: String(c.name)
      }))
    };
  }

  throw new Error('Formato inválido: o arquivo JSON não contém uma lista válida de etiquetas.');
}
