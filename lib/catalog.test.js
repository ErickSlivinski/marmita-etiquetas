import test from 'node:test';
import assert from 'node:assert/strict';
import {
  DEFAULT_CATEGORIES,
  LABELS_PER_SHEET,
  cleanOcrLabelName,
  inferCategoryFromText,
  filterAndSortLabels,
  getCategoryCounts,
  calculateCheckoutSummary,
  createCategorySlug,
  validateBackupPayload
} from './catalog.js';

test('cleanOcrLabelName ignores nutritional metadata and formats title', () => {
  const raw = `
    Peso Líq. 350g
    FRANGO COM BATATA DOCE
    Low Carb
    Ingredientes: peito de frango...
  `;
  assert.equal(cleanOcrLabelName(raw, 'fallback'), 'Frango Com Batata Doce - Low Carb');
});

test('cleanOcrLabelName falls back when OCR text is empty', () => {
  assert.equal(cleanOcrLabelName('   \n  ', 'Minha Etiqueta'), 'Minha Etiqueta');
});

test('inferCategoryFromText maps dishes to matching categories', () => {
  assert.equal(inferCategoryFromText('Frango Desfiado Low Carb'), 'low-carb');
  assert.equal(inferCategoryFromText('Escondidinho de Mandioquinha'), 'massas');
  assert.equal(inferCategoryFromText('Bowl Vegetariano de Grão de Bico'), 'vegetariana');
  assert.equal(inferCategoryFromText('Filé de Frango à Parmegiana'), 'frango');
  assert.equal(inferCategoryFromText('Patinho Moído com Legumes'), 'carnes');
  assert.equal(inferCategoryFromText('Item Genérico 123'), 'sem-categoria');
});

test('filterAndSortLabels filters by category and accent-insensitive search', () => {
  const sample = [
    { id: '1', name: 'Escondidinho de Mandioquinha', categoryId: 'massas', createdAt: 100 },
    { id: '2', name: 'Filé de Frango Low Carb', categoryId: 'low-carb', createdAt: 200 },
    { id: '3', name: 'Abóbora Assada Vegana', categoryId: 'vegetariana', createdAt: 300 }
  ];

  assert.equal(filterAndSortLabels(sample, { categoryId: 'low-carb' }).length, 1);
  assert.equal(filterAndSortLabels(sample, { query: 'file frango' })[0].id, '2');
  assert.equal(filterAndSortLabels(sample, { query: 'abobora' })[0].id, '3');

  const selectedFirst = filterAndSortLabels(sample, {
    sortBy: 'selected-first',
    cart: { '3': 2 }
  });
  assert.equal(selectedFirst[0].id, '3');
});

test('getCategoryCounts counts all labels and per category', () => {
  const sample = [
    { id: '1', categoryId: 'low-carb' },
    { id: '2', categoryId: 'low-carb' },
    { id: '3', categoryId: 'vegetariana' }
  ];
  const counts = getCategoryCounts(sample, DEFAULT_CATEGORIES);
  assert.equal(counts.all, 3);
  assert.equal(counts['low-carb'], 2);
  assert.equal(counts.vegetariana, 1);
  assert.equal(counts.carnes, 0);
});

test('calculateCheckoutSummary computes sheets and total stickers', () => {
  const sample = [
    { id: 'a', name: 'A' },
    { id: 'b', name: 'B' },
    { id: 'c', name: 'C' }
  ];
  const summary = calculateCheckoutSummary(sample, { a: 2, b: 3, c: 0 });
  assert.equal(summary.totalModels, 2);
  assert.equal(summary.totalSheets, 5);
  assert.equal(summary.totalStickers, 5 * LABELS_PER_SHEET);
});

test('createCategorySlug generates unique slugs', () => {
  assert.equal(createCategorySlug('Proteína Extra', DEFAULT_CATEGORIES), 'proteina-extra');
  assert.equal(createCategorySlug('Low Carb', DEFAULT_CATEGORIES), 'low-carb-2');
});

test('validateBackupPayload supports both legacy array and v2 object formats', () => {
  const legacy = validateBackupPayload([{ name: 'Teste', dataUrl: 'data:image/png;base64,abc' }]);
  assert.equal(legacy.labels.length, 1);
  assert.equal(legacy.labels[0].categoryId, 'sem-categoria');

  const v2 = validateBackupPayload({
    categories: [{ id: 'fit', name: 'Fit' }],
    labels: [{ name: 'Frango', dataUrl: 'data:image/png;base64,xyz', categoryId: 'fit' }]
  });
  assert.equal(v2.categories.length, 1);
  assert.equal(v2.labels[0].categoryId, 'fit');

  assert.throws(() => validateBackupPayload({ invalid: true }));
});
