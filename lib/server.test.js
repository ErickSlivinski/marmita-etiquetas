import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from 'pdf-lib';
import { DEFAULT_CATEGORIES } from './catalog.js';
import { buildOrderPdfBuffer } from './pdfServer.js';
import crypto from 'crypto';
import { resolveOrderItems, searchCatalogLabels, readExportPdf, sanitizePdfFilename, verifyMcpRequestAuth } from './serverStore.js';
import { resolvePublicBaseUrl } from './mcpServer.js';
import {
  exchangeOAuthCode,
  isAllowedRedirectUri,
  issueOAuthCodeAndTokensForTest,
  verifyOAuthBearerToken
} from './oauthServer.js';

// Minimal valid 1x1 PNG data URL
const TINY_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

test('buildOrderPdfBuffer generates valid A4 Landscape PDF for mixed 6x3 and 10x4 orders', async () => {
  const items = [
    {
      label: { id: 'l1', name: 'Strogonoff Fit', categoryId: 'low-carb', format: '6x3', dataUrl: TINY_PNG.replace('image/png', 'application/octet-stream') },
      sheets: 2,
      format: '6x3'
    },
    {
      label: { id: 'l2', name: 'Desinberry', categoryId: 'sucos', format: '10x4', dataUrl: TINY_PNG },
      sheets: 1,
      format: '10x4'
    }
  ];

  const buffer = await buildOrderPdfBuffer(items, DEFAULT_CATEGORIES);
  assert.ok(buffer instanceof Buffer);
  assert.equal(buffer.subarray(0, 5).toString('ascii'), '%PDF-');

  const doc = await PDFDocument.load(buffer);
  assert.equal(doc.getPageCount(), 3);
  const { width, height } = doc.getPage(0).getSize();
  // 297 x 210 mm in points (~841.89 x 595.28 pt)
  assert.ok(Math.abs(width - (297 * 72) / 25.4) < 0.5);
  assert.ok(Math.abs(height - (210 * 72) / 25.4) < 0.5);
});

test('resolveOrderItems and searchCatalogLabels match accent-insensitively and merge quantities', () => {
  const catalog = {
    categories: DEFAULT_CATEGORIES,
    labels: [
      { id: 'a1', name: 'Strogonoff Fit De Frango', categoryId: 'low-carb', format: '6x3', dataUrl: TINY_PNG },
      { id: 's1', name: 'Desinberry', categoryId: 'sucos', format: '10x4', dataUrl: TINY_PNG }
    ]
  };

  const hits = searchCatalogLabels(catalog, { query: 'strogonoff', categoryId: 'low-carb' });
  assert.equal(hits.length, 1);
  assert.equal(hits[0].id, 'a1');

  const order = resolveOrderItems(catalog, [
    { nome: 'strogonoff frango', folhas: 2 },
    { etiquetaId: 's1', folhas: 1 }
  ]);
  assert.equal(order.items.length, 2);
  assert.equal(order.summary.totalSheets, 3);
  assert.equal(order.summary.totalStickers, 2 * 24 + 1 * 14);
});

test('sanitizePdfFilename and resolvePublicBaseUrl handle safe defaults cleanly', () => {
  assert.equal(sanitizePdfFilename('../pedido secreto!!.pdf'), 'pedido_secreto.pdf');
  assert.match(sanitizePdfFilename(''), /^etiquetas_pedido_\d{4}-\d{2}-\d{2}\.pdf$/);
  assert.equal(resolvePublicBaseUrl(null), 'http://localhost:3000');
});

test('readExportPdf rejects invalid tokens and path traversal attempts', async () => {
  assert.equal(await readExportPdf('../../package.json'), null);
  assert.equal(await readExportPdf('not-a-valid-token'), null);
});

test('verifyMcpRequestAuth accepts Authorization Bearer, x-api-key, and x-auth-token', async () => {
  const prev = process.env.MCP_API_TOKEN;
  process.env.MCP_API_TOKEN = 'etq_test_secret_12345';
  try {
    const mkReq = (hdrs) => ({ get: (k) => hdrs[k.toLowerCase()] });
    assert.equal(await verifyMcpRequestAuth(mkReq({ authorization: 'Bearer etq_test_secret_12345' })), true);
    assert.equal(await verifyMcpRequestAuth(mkReq({ 'x-api-key': 'etq_test_secret_12345' })), true);
    assert.equal(await verifyMcpRequestAuth(mkReq({ 'x-auth-token': 'etq_test_secret_12345' })), true);
    assert.equal(await verifyMcpRequestAuth(mkReq({ 'x-api-key': 'wrong_token' })), false);
  } finally {
    if (prev === undefined) delete process.env.MCP_API_TOKEN;
    else process.env.MCP_API_TOKEN = prev;
  }
});

test('OAuth PKCE flow validates allowlisted redirects, exchanges single-use codes, and verifies access tokens', () => {
  assert.equal(isAllowedRedirectUri('https://claude.ai/api/mcp/auth_callback'), true);
  assert.equal(isAllowedRedirectUri('https://claude.com/api/mcp/auth_callback'), true);
  assert.equal(isAllowedRedirectUri('http://127.0.0.1:45123/oauth/callback'), true);
  assert.equal(isAllowedRedirectUri('https://evil.example.com/callback'), false);
  assert.equal(isAllowedRedirectUri('http://claude.ai/callback'), false);

  const verifier = 'a_very_long_pkce_code_verifier_string_with_enough_entropy_123456';
  const challenge = crypto
    .createHash('sha256')
    .update(verifier)
    .digest('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');

  const { code } = issueOAuthCodeAndTokensForTest({
    clientId: 'etq_cli_test',
    redirectUri: 'https://claude.ai/api/mcp/auth_callback',
    codeChallenge: challenge
  });

  const badVerifier = exchangeOAuthCode({
    code,
    clientId: 'etq_cli_test',
    redirectUri: 'https://claude.ai/api/mcp/auth_callback',
    codeVerifier: 'wrong_verifier_string_with_43_characters_minimum_padding_000'
  });
  assert.equal(badVerifier.error, 'invalid_grant');

  const ok = exchangeOAuthCode({
    code,
    clientId: 'etq_cli_test',
    redirectUri: 'https://claude.ai/api/mcp/auth_callback',
    codeVerifier: verifier
  });
  assert.ok(ok.access_token);
  assert.equal(verifyOAuthBearerToken(`Bearer ${ok.access_token}`), true);

  // Code replay must fail
  const replay = exchangeOAuthCode({
    code,
    clientId: 'etq_cli_test',
    redirectUri: 'https://claude.ai/api/mcp/auth_callback',
    codeVerifier: verifier
  });
  assert.equal(replay.error, 'invalid_grant');
});

test('resolvePublicBaseUrl strips accidental /mcp suffix from PUBLIC_APP_URL', () => {
  const prev = process.env.PUBLIC_APP_URL;
  process.env.PUBLIC_APP_URL = 'https://marmita-etiquetas.onrender.com/mcp';
  try {
    assert.equal(resolvePublicBaseUrl(null), 'https://marmita-etiquetas.onrender.com');
  } finally {
    if (prev === undefined) delete process.env.PUBLIC_APP_URL;
    else process.env.PUBLIC_APP_URL = prev;
  }
});
