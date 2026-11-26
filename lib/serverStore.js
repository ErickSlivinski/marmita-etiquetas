import fs from 'fs/promises';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import {
  DEFAULT_CATEGORIES,
  LABEL_FORMATS,
  clampSheets,
  normalizeText,
  resolveLabelFormat,
  validateBackupPayload
} from './catalog.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const ROOT_DIR = path.resolve(__dirname, '..');
const DATA_DIR = path.resolve(ROOT_DIR, 'data');
const EXPORTS_DIR = path.resolve(DATA_DIR, 'exports');
const CATALOG_PATH = path.resolve(DATA_DIR, 'catalog.json');
const AUTH_PATH = path.resolve(DATA_DIR, 'auth.json');
const SEED_BACKUP_PATH = path.resolve(ROOT_DIR, 'public', 'etiquetafit_backup_atualizado.json');

const EXPORT_TTL_MS = 24 * 60 * 60 * 1000;
const MAX_EXPORTS_ON_DISK = 50;

let catalogCache = null;

async function ensureDataDirs() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.mkdir(EXPORTS_DIR, { recursive: true });
}

function sha256Buffer(value) {
  return crypto.createHash('sha256').update(String(value || '')).digest();
}

function constantTimeMatch(a, b) {
  if (!a || !b) return false;
  return crypto.timingSafeEqual(sha256Buffer(a), sha256Buffer(b));
}

function buildStableSeedId(label, idx) {
  if (label.id && typeof label.id === 'string') return label.id;
  const hash = crypto
    .createHash('sha1')
    .update(`${label.name}:${label.dataUrl.slice(0, 64)}:${idx}`)
    .digest('hex')
    .slice(0, 12);
  return `lbl-${hash}`;
}

function normalizeCategories(rawCategories) {
  const base = Array.isArray(rawCategories) && rawCategories.length > 0 ? rawCategories : [...DEFAULT_CATEGORIES];
  return base.map(c => ({
    id: String(c.id),
    name: String(c.name),
    defaultFormat: LABEL_FORMATS[c.defaultFormat]
      ? c.defaultFormat
      : (/\bsucos?\b/.test(normalizeText(c.name)) || c.id === 'sucos' ? '10x4' : '6x3')
  }));
}

function normalizeLabels(rawLabels, categories) {
  return rawLabels.map((l, idx) => ({
    id: buildStableSeedId(l, idx),
    name: String(l.name),
    categoryId: categories.some(c => c.id === l.categoryId) ? l.categoryId : 'sem-categoria',
    format: LABEL_FORMATS[l.format] ? l.format : undefined,
    dataUrl: String(l.dataUrl),
    createdAt: Number(l.createdAt) || Date.now()
  }));
}

export async function getCatalog() {
  if (catalogCache) return catalogCache;

  await ensureDataDirs();

  try {
    const raw = await fs.readFile(CATALOG_PATH, 'utf8');
    const parsed = JSON.parse(raw);
    const categories = normalizeCategories(parsed.categories);
    const labels = normalizeLabels(parsed.labels || [], categories);
    catalogCache = {
      categories,
      labels,
      syncedAt: parsed.syncedAt || new Date().toISOString(),
      syncSource: parsed.syncSource || 'server'
    };
    return catalogCache;
  } catch {
    // Seed from updated backup JSON if present
    try {
      const seedRaw = await fs.readFile(SEED_BACKUP_PATH, 'utf8');
      const seedParsed = validateBackupPayload(JSON.parse(seedRaw));
      const categories = normalizeCategories(seedParsed.categories);
      const labels = normalizeLabels(seedParsed.labels, categories);
      catalogCache = {
        categories,
        labels,
        syncedAt: new Date().toISOString(),
        syncSource: 'backup-seed'
      };
      await fs.writeFile(CATALOG_PATH, JSON.stringify(catalogCache), 'utf8');
      return catalogCache;
    } catch {
      const categories = normalizeCategories(DEFAULT_CATEGORIES);
      catalogCache = {
        categories,
        labels: [],
        syncedAt: null,
        syncSource: 'empty'
      };
      return catalogCache;
    }
  }
}

export async function syncCatalog(payload, syncSource = 'browser-sync') {
  const validated = validateBackupPayload(payload);
  const categories = normalizeCategories(validated.categories);
  const labels = normalizeLabels(
    validated.labels.map((l, idx) => ({
      ...l,
      id: payload.labels?.[idx]?.id || l.id
    })),
    categories
  );

  const nextState = {
    categories,
    labels,
    syncedAt: new Date().toISOString(),
    syncSource
  };

  await ensureDataDirs();
  await fs.writeFile(CATALOG_PATH, JSON.stringify(nextState), 'utf8');
  catalogCache = nextState;

  return {
    labelCount: labels.length,
    categoryCount: categories.length,
    syncedAt: nextState.syncedAt,
    syncSource: nextState.syncSource
  };
}

export function searchCatalogLabels(catalog, { query = '', categoryId = 'all', limit = 30 } = {}) {
  const { labels, categories } = catalog;
  const catMap = new Map(categories.map(c => [c.id, c.name]));
  const normQuery = normalizeText(query);
  const terms = normQuery ? normQuery.split(' ').filter(Boolean) : [];
  const normCatFilter = categoryId && categoryId !== 'all' ? normalizeText(categoryId) : null;

  const matches = [];
  for (const label of labels) {
    const catName = catMap.get(label.categoryId) || 'Sem Categoria';
    if (normCatFilter) {
      const matchesCatId = normalizeText(label.categoryId) === normCatFilter;
      const matchesCatName = normalizeText(catName).includes(normCatFilter);
      if (!matchesCatId && !matchesCatName) continue;
    }

    if (terms.length > 0) {
      const haystack = normalizeText(`${label.name} ${catName}`);
      if (!terms.every(t => haystack.includes(t))) continue;
    }

    const format = resolveLabelFormat(label, categories);
    matches.push({
      id: label.id,
      name: label.name,
      categoryId: label.categoryId,
      categoryName: catName,
      format,
      labelsPerSheet: LABEL_FORMATS[format].labelsPerSheet
    });
  }

  matches.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'base' }));
  return matches.slice(0, Math.min(Math.max(1, Number(limit) || 30), 200));
}

function findBestLabelMatch(catalog, identifier) {
  const { labels, categories } = catalog;
  if (!identifier) return { match: null, candidates: [] };

  const exactId = labels.find(l => l.id === identifier);
  if (exactId) return { match: exactId, candidates: [] };

  const normNeedle = normalizeText(identifier);
  if (!normNeedle) return { match: null, candidates: [] };

  const exactNorm = labels.find(l => normalizeText(l.name) === normNeedle);
  if (exactNorm) return { match: exactNorm, candidates: [] };

  const needleTerms = normNeedle.split(' ').filter(Boolean);
  const catMap = new Map(categories.map(c => [c.id, c.name]));

  const scored = [];
  for (const label of labels) {
    const normName = normalizeText(label.name);
    const haystack = normalizeText(`${label.name} ${catMap.get(label.categoryId) || ''}`);
    let score = 0;

    if (normName.startsWith(normNeedle)) score += 100;
    else if (normName.includes(normNeedle)) score += 75;
    else if (needleTerms.every(t => normName.includes(t))) score += 60;
    else if (needleTerms.every(t => haystack.includes(t))) score += 40;
    else {
      const hitCount = needleTerms.filter(t => normName.includes(t)).length;
      if (hitCount > 0 && hitCount >= Math.ceil(needleTerms.length * 0.6)) {
        score += hitCount * 12;
      }
    }

    if (score > 0) {
      scored.push({ label, score });
    }
  }

  scored.sort((a, b) => b.score - a.score || a.label.name.length - b.label.name.length);
  if (scored.length === 0) return { match: null, candidates: [] };

  return {
    match: scored[0].label,
    candidates: scored.slice(0, 5).map(s => s.label.name)
  };
}

export function resolveOrderItems(catalog, requestedItems) {
  if (!Array.isArray(requestedItems) || requestedItems.length === 0) {
    throw new Error('Informe ao menos 1 etiqueta e a quantidade de folhas.');
  }

  const resolvedMap = new Map();
  const unresolved = [];

  for (const reqItem of requestedItems) {
    if (!reqItem || typeof reqItem !== 'object') continue;
    const rawSheets = Number(reqItem.sheets ?? reqItem.folhas ?? 1);
    const sheets = clampSheets(rawSheets);
    if (sheets <= 0) continue;

    const identifier = reqItem.labelId || reqItem.etiquetaId || reqItem.name || reqItem.nome || reqItem.query || '';
    const { match } = findBestLabelMatch(catalog, identifier);

    if (!match) {
      unresolved.push(String(identifier || 'Item sem nome'));
      continue;
    }

    const format = resolveLabelFormat(match, catalog.categories);
    const perSheet = LABEL_FORMATS[format].labelsPerSheet;

    if (resolvedMap.has(match.id)) {
      const prev = resolvedMap.get(match.id);
      const mergedSheets = clampSheets(prev.sheets + sheets);
      prev.sheets = mergedSheets;
      prev.totalStickers = mergedSheets * perSheet;
    } else {
      resolvedMap.set(match.id, {
        label: match,
        sheets,
        format,
        labelsPerSheet: perSheet,
        totalStickers: sheets * perSheet
      });
    }
  }

  const items = Array.from(resolvedMap.values());
  let totalSheets = 0;
  let totalStickers = 0;
  const sheetsByFormat = { '6x3': 0, '10x4': 0 };

  for (const item of items) {
    totalSheets += item.sheets;
    totalStickers += item.totalStickers;
    sheetsByFormat[item.format] = (sheetsByFormat[item.format] || 0) + item.sheets;
  }

  return {
    items,
    unresolved,
    summary: {
      totalModels: items.length,
      totalSheets,
      totalStickers,
      sheetsByFormat
    }
  };
}

// ─── PDF Export Storage (with strict token confinement) ───────────────────────

export function sanitizePdfFilename(rawName) {
  const cleaned = String(rawName || '')
    .replace(/\.pdf$/i, '')
    .replace(/[^a-zA-Z0-9À-ÿ_\-\s]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .slice(0, 64);
  const date = new Date().toISOString().split('T')[0];
  return `${cleaned || `etiquetas_pedido_${date}`}.pdf`;
}

export async function saveExportPdf(buffer, rawFilename) {
  await ensureDataDirs();
  const token = crypto.randomBytes(16).toString('hex'); // 32 hex chars
  const filename = sanitizePdfFilename(rawFilename);
  const targetPath = path.resolve(EXPORTS_DIR, `${token}.pdf`);
  if (!targetPath.startsWith(EXPORTS_DIR + path.sep)) {
    throw new Error('Caminho de exportação inválido.');
  }

  await fs.writeFile(targetPath, buffer);
  await pruneOldExports().catch(() => {});

  return { token, filename };
}

export async function readExportPdf(token) {
  if (typeof token !== 'string' || !/^[a-f0-9]{32}$/.test(token)) {
    return null;
  }
  await ensureDataDirs();
  const targetPath = path.resolve(EXPORTS_DIR, `${token}.pdf`);
  if (!targetPath.startsWith(EXPORTS_DIR + path.sep)) {
    return null;
  }
  try {
    const stat = await fs.stat(targetPath);
    if (Date.now() - stat.mtimeMs > EXPORT_TTL_MS) {
      await fs.unlink(targetPath).catch(() => {});
      return null;
    }
    return await fs.readFile(targetPath);
  } catch {
    return null;
  }
}

async function pruneOldExports() {
  const entries = await fs.readdir(EXPORTS_DIR, { withFileTypes: true });
  const files = [];
  const now = Date.now();
  for (const entry of entries) {
    if (!entry.isFile() || !/^[a-f0-9]{32}\.pdf$/.test(entry.name)) continue;
    const full = path.resolve(EXPORTS_DIR, entry.name);
    const st = await fs.stat(full).catch(() => null);
    if (!st) continue;
    if (now - st.mtimeMs > EXPORT_TTL_MS) {
      await fs.unlink(full).catch(() => {});
    } else {
      files.push({ full, mtimeMs: st.mtimeMs });
    }
  }
  if (files.length > MAX_EXPORTS_ON_DISK) {
    files.sort((a, b) => a.mtimeMs - b.mtimeMs);
    for (const old of files.slice(0, files.length - MAX_EXPORTS_ON_DISK)) {
      await fs.unlink(old.full).catch(() => {});
    }
  }
}

// ─── Auth Token & Non-Secret Settings ────────────────────────────────────────

async function readAuthRecord() {
  try {
    const raw = await fs.readFile(AUTH_PATH, 'utf8');
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export async function getAuthStatus() {
  const envToken = (process.env.MCP_API_TOKEN || '').trim();
  if (envToken) {
    const preview = envToken.length >= 8
      ? `${envToken.slice(0, 4)}****${envToken.slice(-4)}`
      : '****';
    return {
      configured: true,
      source: 'env',
      preview
    };
  }

  const record = await readAuthRecord();
  if (record && record.hash && record.salt) {
    return {
      configured: true,
      source: 'generated',
      preview: record.preview || 'etq_****'
    };
  }

  return {
    configured: false,
    source: 'none',
    preview: null
  };
}

export async function claimOrRotateMcpToken() {
  const envToken = (process.env.MCP_API_TOKEN || '').trim();
  if (envToken) {
    throw new Error('O token MCP está definido via variável de ambiente MCP_API_TOKEN no servidor.');
  }

  await ensureDataDirs();
  const plainToken = `etq_${crypto.randomBytes(20).toString('hex')}`;
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(plainToken, salt, 64).toString('hex');
  const preview = `${plainToken.slice(0, 6)}****${plainToken.slice(-4)}`;

  await fs.writeFile(
    AUTH_PATH,
    JSON.stringify({ salt, hash, preview, updatedAt: new Date().toISOString() }),
    'utf8'
  );

  return { token: plainToken, preview };
}

export async function verifyRawMcpToken(provided) {
  if (typeof provided !== 'string') return false;
  const clean = provided.trim();
  if (!clean) return false;

  const envToken = (process.env.MCP_API_TOKEN || '').trim();
  if (envToken) {
    return constantTimeMatch(clean, envToken);
  }

  const record = await readAuthRecord();
  if (!record || !record.salt || !record.hash) {
    return false;
  }

  const derived = crypto.scryptSync(clean, record.salt, 64);
  const stored = Buffer.from(record.hash, 'hex');
  if (derived.length !== stored.length) return false;
  return crypto.timingSafeEqual(derived, stored);
}

export async function verifyMcpBearerToken(authHeader) {
  if (typeof authHeader !== 'string') return false;
  const match = authHeader.match(/^Bearer\s+(.+)$/i);
  if (!match) return false;
  return verifyRawMcpToken(match[1]);
}

export async function verifyMcpRequestAuth(req) {
  const authHeader = req.get('authorization');
  if (authHeader && (await verifyMcpBearerToken(authHeader))) {
    return true;
  }

  const altHeader = req.get('x-api-key') || req.get('x-auth-token');
  if (typeof altHeader === 'string' && altHeader.trim()) {
    const stripped = altHeader.trim().replace(/^Bearer\s+/i, '');
    return verifyRawMcpToken(stripped);
  }

  return false;
}

export function isTrustedUiRequest(req) {
  const customHeader = req.get('x-requested-with');
  if (customHeader !== 'EtiquetaFit') return false;

  const allowedHosts = new Set(
    [req.get('x-forwarded-host'), req.get('host')]
      .filter(Boolean)
      .flatMap(v => String(v).split(','))
      .map(v => v.trim().toLowerCase())
      .filter(Boolean)
  );
  if (allowedHosts.size === 0) return false;

  const origin = req.get('origin');
  const referer = req.get('referer');

  const checkHostMatches = (urlStr) => {
    try {
      const parsed = new URL(urlStr);
      return allowedHosts.has(parsed.host.toLowerCase());
    } catch {
      return false;
    }
  };

  if (origin && !checkHostMatches(origin)) return false;
  if (!origin && referer && !checkHostMatches(referer)) return false;
  return true;
}
