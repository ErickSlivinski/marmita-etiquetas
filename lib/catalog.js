/**
 * Pure domain logic for EtiquetaFit Dashboard
 * Shared between browser UI (app.js) and Node unit tests (lib/catalog.test.js).
 */

export const LABELS_PER_SHEET = 24;
export const JUICE_LABELS_PER_SHEET = 14;
export const MAX_SHEETS_PER_LABEL = 50;

export const LABEL_FORMATS = {
  '6x3': {
    id: '6x3',
    name: 'Comum (6×3 cm)',
    shortDim: '6×3 cm',
    labelsPerSheet: 24,
    orientation: 'landscape',
    pageW: 297,
    pageH: 210,
    cols: 4,
    rows: 6,
    marginX: 5,
    marginY: 5,
    gapX: 0,
    gapY: 0,
    labelW: (297 - 10) / 4,
    labelH: (210 - 10) / 6
  },
  '10x4': {
    id: '10x4',
    name: 'Sucos (10×4 cm)',
    shortDim: '10×4 cm',
    labelsPerSheet: 14,
    orientation: 'landscape',
    rotate90: true,
    pageW: 297,
    pageH: 210,
    cols: 7,
    rows: 2,
    marginX: 8.5,
    marginY: 5,
    gapX: 0,
    gapY: 0,
    labelW: 40,
    labelH: 100
  }
};

export const DEFAULT_CATEGORIES = [
  { id: 'low-carb', name: 'Low Carb', defaultFormat: '6x3' },
  { id: 'vegetariana', name: 'Vegetariana', defaultFormat: '6x3' },
  { id: 'frango', name: 'Frango & Aves', defaultFormat: '6x3' },
  { id: 'carnes', name: 'Carnes', defaultFormat: '6x3' },
  { id: 'massas', name: 'Massas & Escondidinhos', defaultFormat: '6x3' },
  { id: 'sucos', name: 'Sucos', defaultFormat: '10x4' },
  { id: 'sem-categoria', name: 'Sem Categoria', defaultFormat: '6x3' }
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

export function isSucosCategory(categoryId, categories = DEFAULT_CATEGORIES) {
  if (!categoryId) return false;
  const normId = normalizeText(categoryId);
  if (normId === 'suco' || normId === 'sucos') return true;
  const cat = categories.find(c => c.id === categoryId);
  if (!cat) return false;
  if (cat.defaultFormat === '10x4') return true;
  return /\bsucos?\b/.test(normalizeText(cat.name));
}

export function findSucosCategoryId(categories = DEFAULT_CATEGORIES) {
  const byExact = categories.find(c => c.id === 'sucos' || c.id === 'suco');
  if (byExact) return byExact.id;
  const byFormat = categories.find(c => c.defaultFormat === '10x4');
  if (byFormat) return byFormat.id;
  const byName = categories.find(c => /\bsucos?\b/.test(normalizeText(c.name)));
  return byName ? byName.id : null;
}

export function detectFormatFromDimensions(width, height) {
  if (!width || !height || width <= 0 || height <= 0) return '6x3';
  const ratio = width / height;
  // Standard meal labels are 6x3 cm (ratio 2.0); juice labels are 10x4 cm (ratio 2.5)
  return ratio >= 2.25 ? '10x4' : '6x3';
}

export function resolveLabelFormat(label, categories = DEFAULT_CATEGORIES) {
  if (label?.format && LABEL_FORMATS[label.format]) {
    return label.format;
  }
  if (isSucosCategory(label?.categoryId, categories)) {
    return '10x4';
  }
  return '6x3';
}

export function getLabelFormatSpec(formatOrLabel, categories = DEFAULT_CATEGORIES) {
  if (typeof formatOrLabel === 'string') {
    return LABEL_FORMATS[formatOrLabel] || LABEL_FORMATS['6x3'];
  }
  const fmt = resolveLabelFormat(formatOrLabel, categories);
  return LABEL_FORMATS[fmt] || LABEL_FORMATS['6x3'];
}

export function cleanOcrLabelName(rawText, fallbackName = 'Etiqueta sem nome') {
  if (!rawText || typeof rawText !== 'string') return fallbackName;

  const ignoredPatterns = [
    /^peso\s*l[ií]q/i,
    /^val[:\s]/i,
    /^validade/i,
    /^fab[:\s]/i,
    /^fabricac[aã]o/i,
    /^lote/i,
    /^ingredientes/i,
    /^informa[cç][aã]o\s+nutricional/i,
    /^modo\s+de\s+preparo/i,
    /^conservar/i,
    /^@/,
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

export function inferCategoryFromText(text, categories = DEFAULT_CATEGORIES, detectedFormat = null) {
  if (detectedFormat === '10x4') {
    const sucosId = findSucosCategoryId(categories);
    if (sucosId) return sucosId;
  }

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

  if (/\b(suco|sucos|desinberry|detox|hibisco|cha de hibisco|smoothie|limonada|polpa|vitamina)\b/.test(norm)) {
    const sucosId = findSucosCategoryId(categories);
    if (sucosId) return sucosId;
  }

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

export function calculateCheckoutSummary(labels, cart, categoriesOrOverride = DEFAULT_CATEGORIES) {
  const isFixedNumber = typeof categoriesOrOverride === 'number';
  const categories = isFixedNumber ? DEFAULT_CATEGORIES : (categoriesOrOverride || DEFAULT_CATEGORIES);

  const items = [];
  let totalSheets = 0;
  let totalStickers = 0;
  const sheetsByFormat = { '6x3': 0, '10x4': 0 };

  for (const label of labels) {
    const sheets = clampSheets(cart[label.id]);
    if (sheets > 0) {
      const format = resolveLabelFormat(label, categories);
      const perSheet = isFixedNumber
        ? categoriesOrOverride
        : LABEL_FORMATS[format].labelsPerSheet;
      const itemStickers = sheets * perSheet;

      items.push({
        label,
        sheets,
        format,
        labelsPerSheet: perSheet,
        totalStickers: itemStickers
      });

      totalSheets += sheets;
      totalStickers += itemStickers;
      sheetsByFormat[format] = (sheetsByFormat[format] || 0) + sheets;
    }
  }

  return {
    items,
    totalModels: items.length,
    totalSheets,
    totalStickers,
    sheetsByFormat
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
        format: LABEL_FORMATS[l.format] ? l.format : undefined,
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
        format: LABEL_FORMATS[l.format] ? l.format : undefined,
        createdAt: l.createdAt || Date.now()
      })),
      categories: validCategories.map(c => ({
        id: String(c.id),
        name: String(c.name),
        defaultFormat: LABEL_FORMATS[c.defaultFormat]
          ? c.defaultFormat
          : (/\bsucos?\b/.test(normalizeText(c.name)) || c.id === 'sucos' ? '10x4' : '6x3')
      }))
    };
  }

  throw new Error('Formato inválido: o arquivo JSON não contém uma lista válida de etiquetas.');
}

