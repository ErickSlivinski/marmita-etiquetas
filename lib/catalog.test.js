import test from 'node:test';
import assert from 'node:assert/strict';
import {
  DEFAULT_CATEGORIES,
  LABELS_PER_SHEET,
  JUICE_LABELS_PER_SHEET,
  LABEL_FORMATS,
  cleanOcrLabelName,
  detectFormatFromDimensions,
  resolveLabelFormat,
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

test('detectFormatFromDimensions distinguishes 6x3 and 10x4 labels by aspect ratio', () => {
  assert.equal(detectFormatFromDimensions(600, 300), '6x3');
  assert.equal(detectFormatFromDimensions(709, 354), '6x3');
  assert.equal(detectFormatFromDimensions(1181, 472), '10x4');
  assert.equal(detectFormatFromDimensions(1000, 400), '10x4');
  assert.equal(detectFormatFromDimensions(0, 0), '6x3');
});

test('resolveLabelFormat uses explicit format or infers 10x4 from Sucos category', () => {
  assert.equal(resolveLabelFormat({ categoryId: 'low-carb' }, DEFAULT_CATEGORIES), '6x3');
  assert.equal(resolveLabelFormat({ categoryId: 'sucos' }, DEFAULT_CATEGORIES), '10x4');
  assert.equal(
    resolveLabelFormat({ categoryId: 'bebidas' }, [{ id: 'bebidas', name: 'Sucos Naturais' }]),
    '10x4'
  );
  assert.equal(
    resolveLabelFormat({ categoryId: 'low-carb', format: '10x4' }, DEFAULT_CATEGORIES),
    '10x4'
  );
  assert.equal(
    resolveLabelFormat({ categoryId: 'sucos', format: '6x3' }, DEFAULT_CATEGORIES),
    '6x3'
  );
});

test('inferCategoryFromText maps dishes and juices to matching categories', () => {
  assert.equal(inferCategoryFromText('Frango Desfiado Low Carb'), 'low-carb');
  assert.equal(inferCategoryFromText('Escondidinho de Mandioquinha'), 'massas');
  assert.equal(inferCategoryFromText('Bowl Vegetariano de Grão de Bico'), 'vegetariana');
  assert.equal(inferCategoryFromText('Filé de Frango à Parmegiana'), 'frango');
  assert.equal(inferCategoryFromText('Patinho Moído com Legumes'), 'carnes');
  assert.equal(inferCategoryFromText('DESINBERRY Morango, abacaxi e chá de hibisco'), 'sucos');
  assert.equal(inferCategoryFromText('Qualquer texto', DEFAULT_CATEGORIES, '10x4'), 'sucos');
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

test('calculateCheckoutSummary computes mixed 6x3 (24/sheet) and 10x4 (14/sheet) totals', () => {
  const sample = [
    { id: 'a', name: 'Marmita A', categoryId: 'low-carb' },
    { id: 'b', name: 'Marmita B', categoryId: 'carnes' },
    { id: 'j', name: 'Suco Amarelo', categoryId: 'sucos', format: '10x4' },
    { id: 'c', name: 'Marmita C', categoryId: 'massas' }
  ];

  const stdOnly = calculateCheckoutSummary(sample, { a: 2, b: 3, c: 0 });
  assert.equal(stdOnly.totalModels, 2);
  assert.equal(stdOnly.totalSheets, 5);
  assert.equal(stdOnly.totalStickers, 5 * LABELS_PER_SHEET);

  const mixed = calculateCheckoutSummary(sample, { a: 1, j: 2 });
  assert.equal(mixed.totalModels, 2);
  assert.equal(mixed.totalSheets, 3);
  assert.equal(mixed.totalStickers, 1 * LABELS_PER_SHEET + 2 * JUICE_LABELS_PER_SHEET);
  assert.equal(mixed.sheetsByFormat['6x3'], 1);
  assert.equal(mixed.sheetsByFormat['10x4'], 2);
  assert.equal(LABEL_FORMATS['10x4'].cols * LABEL_FORMATS['10x4'].rows, 14);
});

test('createCategorySlug generates unique slugs', () => {
  assert.equal(createCategorySlug('Proteína Extra', DEFAULT_CATEGORIES), 'proteina-extra');
  assert.equal(createCategorySlug('Low Carb', DEFAULT_CATEGORIES), 'low-carb-2');
});

test('validateBackupPayload supports both legacy array and v2 object formats with format metadata', () => {
  const legacy = validateBackupPayload([{ name: 'Teste', dataUrl: 'data:image/png;base64,abc' }]);
  assert.equal(legacy.labels.length, 1);
  assert.equal(legacy.labels[0].categoryId, 'sem-categoria');

  const v2 = validateBackupPayload({
    categories: [
      { id: 'fit', name: 'Fit' },
      { id: 'sucos', name: 'Sucos' }
    ],
    labels: [
      { name: 'Frango', dataUrl: 'data:image/png;base64,xyz', categoryId: 'fit' },
      { name: 'Desinberry', dataUrl: 'data:image/png;base64,jkl', categoryId: 'sucos', format: '10x4' }
    ]
  });
  assert.equal(v2.categories.length, 2);
  assert.equal(v2.categories[1].defaultFormat, '10x4');
  assert.equal(v2.labels[1].format, '10x4');

  assert.throws(() => validateBackupPayload({ invalid: true }));
});

