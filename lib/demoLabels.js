/**
 * Generates crisp 60x30mm frozen meal and 100x40mm juice label sample PNGs on HTML5 Canvas
 * so users can immediately explore the dashboard catalog and mixed PDF checkout.
 */

const SAMPLE_PRESETS = [
  {
    name: 'Frango Desfiado com Abobrinha',
    subtitle: 'Low Carb · 320g · 285 kcal',
    categoryId: 'low-carb',
    format: '6x3',
    accent: '#0F766E',
    ingredients: 'Peito de frango, abobrinha, azeite, alho e ervas'
  },
  {
    name: 'Escondidinho de Mandioquinha',
    subtitle: 'Clássica · 350g · 410 kcal',
    categoryId: 'massas',
    format: '6x3',
    accent: '#B45309',
    ingredients: 'Mandioquinha, carne moída magra, requeijão light'
  },
  {
    name: 'Bowl Vegetariano de Grão-de-Bico',
    subtitle: 'Vegetariana · 330g · 340 kcal',
    categoryId: 'vegetariana',
    format: '6x3',
    accent: '#15803D',
    ingredients: 'Grão-de-bico, quinoa, abóbora assada e espinafre'
  },
  {
    name: 'Filé de Frango à Parmegiana Fit',
    subtitle: 'Frango & Aves · 350g · 390 kcal',
    categoryId: 'frango',
    format: '6x3',
    accent: '#C2410C',
    ingredients: 'Filé de frango grelhado, molho de tomate natural, muçarela'
  },
  {
    name: 'Patinho Moído com Couve-Flor',
    subtitle: 'Low Carb · 320g · 310 kcal',
    categoryId: 'low-carb',
    format: '6x3',
    accent: '#0F766E',
    ingredients: 'Patinho moído, arroz de couve-flor, cenoura e salsinha'
  },
  {
    name: 'Ragu de Carne com Purê Rústico',
    subtitle: 'Carnes · 360g · 435 kcal',
    categoryId: 'carnes',
    format: '6x3',
    accent: '#991B1B',
    ingredients: 'Acém cozido lentamente, purê de batata inglesa, tomilho'
  },
  {
    name: 'Suco Amarelo',
    subtitle: 'Manga, maçã, limão e hortelã',
    categoryId: 'sucos',
    format: '10x4',
    accent: '#78B62B',
    ingredients: 'Validade: 3 meses no freezer · 10×4 cm'
  },
  {
    name: 'Desinberry',
    subtitle: 'Morango, abacaxi, hortelã, gengibre e chá de hibisco',
    categoryId: 'sucos',
    format: '10x4',
    accent: '#78B62B',
    ingredients: 'Validade: 3 meses no freezer · 10×4 cm'
  }
];

function renderJuiceLabelDataUrl(preset) {
  const canvas = document.createElement('canvas');
  canvas.width = 1000;
  canvas.height = 400;
  const ctx = canvas.getContext('2d');

  // Green outer frame (10x4 cm style)
  ctx.fillStyle = preset.accent;
  ctx.fillRect(0, 0, 1000, 400);

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(20, 20, 960, 360);

  // Leaf/badge emblem on left
  ctx.fillStyle = preset.accent;
  ctx.beginPath();
  ctx.arc(110, 120, 56, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = '#FFFFFF';
  ctx.font = 'bold 26px sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('FIT', 110, 129);

  // Juice Title
  ctx.fillStyle = preset.accent;
  ctx.font = 'bold 62px sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText(preset.name.toUpperCase(), 540, 140, 700);

  // Ingredients
  ctx.fillStyle = '#0F172A';
  ctx.font = '600 28px sans-serif';
  ctx.fillText(preset.subtitle, 540, 225, 760);

  // Footer
  ctx.fillStyle = '#334155';
  ctx.font = '600 22px sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText('@etiquetafit', 56, 342);
  ctx.textAlign = 'right';
  ctx.fillText(preset.ingredients, 944, 342);

  return canvas.toDataURL('image/png');
}

function renderSampleLabelDataUrl(preset) {
  if (preset.format === '10x4') {
    return renderJuiceLabelDataUrl(preset);
  }

  const canvas = document.createElement('canvas');
  canvas.width = 600;
  canvas.height = 300;
  const ctx = canvas.getContext('2d');

  // Crisp sticker paper background
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, 600, 300);

  // Outer cut border
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 4;
  ctx.strokeRect(10, 10, 580, 280);

  // Top color bar
  ctx.fillStyle = preset.accent;
  ctx.fillRect(10, 10, 580, 36);

  // Brand inside sticker
  ctx.fillStyle = '#FFFFFF';
  ctx.font = 'bold 16px sans-serif';
  ctx.fillText('MARMITARIA FIT · CONSERVAR CONGELADO (-18°C)', 26, 34);

  // Dish Title
  ctx.fillStyle = '#0F172A';
  ctx.font = 'bold 28px sans-serif';
  ctx.fillText(preset.name, 26, 98, 548);

  // Subtitle / Nutrition
  ctx.fillStyle = preset.accent;
  ctx.font = '600 18px sans-serif';
  ctx.fillText(preset.subtitle, 26, 134, 548);

  // Divider
  ctx.strokeStyle = '#E2E8F0';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(26, 156);
  ctx.lineTo(574, 156);
  ctx.stroke();

  // Ingredients
  ctx.fillStyle = '#475569';
  ctx.font = '16px sans-serif';
  ctx.fillText('Ingredientes: ' + preset.ingredients, 26, 190, 548);

  // Footer instructions
  ctx.fillStyle = '#64748B';
  ctx.font = '14px sans-serif';
  ctx.fillText('Micro-ondas: 5 a 6 min  |  Validade: 90 dias  |  60×30 mm', 26, 258, 548);

  return canvas.toDataURL('image/png');
}

export function generateDemoLabels() {
  const now = Date.now();
  return SAMPLE_PRESETS.map((preset, index) => ({
    id: crypto.randomUUID(),
    name: preset.name,
    categoryId: preset.categoryId,
    format: preset.format,
    dataUrl: renderSampleLabelDataUrl(preset),
    createdAt: now - index * 1000
  }));
}

