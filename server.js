import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import {
  claimOrRotateMcpToken,
  getAuthStatus,
  getCatalog,
  isTrustedUiRequest,
  readExportPdf,
  sanitizePdfFilename,
  searchCatalogLabels,
  syncCatalog,
  verifyMcpRequestAuth
} from './lib/serverStore.js';
import { createEtiquetaFitMcpServer, createOrderPdfAndDeliver, resolvePublicBaseUrl } from './lib/mcpServer.js';
import { registerOAuthRoutes, verifyOAuthBearerToken } from './lib/oauthServer.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.set('trust proxy', 1);

const PORT = Number(process.env.PORT) || 3000;
const HOST = '0.0.0.0';

app.use(express.json({ limit: '35mb' }));

async function requireApiOrUiAuth(req, res, next) {
  if (isTrustedUiRequest(req)) return next();
  const ok = (await verifyMcpRequestAuth(req)) || verifyOAuthBearerToken(req.get('authorization'));
  if (!ok) {
    return res.status(401).json({ error: 'Não autorizado. Envie Authorization: Bearer <MCP_API_TOKEN> ou x-api-key.' });
  }
  return next();
}

function requireTrustedUi(req, res, next) {
  if (!isTrustedUiRequest(req)) {
    return res.status(403).json({ error: 'Acesso permitido apenas pelo painel EtiquetaFit.' });
  }
  return next();
}

async function requireBearerAuth(req, res, next) {
  const ok = (await verifyMcpRequestAuth(req)) || verifyOAuthBearerToken(req.get('authorization'));
  if (!ok) {
    const baseUrl = resolvePublicBaseUrl(req);
    res.setHeader('WWW-Authenticate', `Bearer resource_metadata="${baseUrl}/.well-known/oauth-protected-resource"`);
    return res.status(401).json({ error: 'Token de autenticação inválido ou ausente.' });
  }
  return next();
}

// MCP OAuth 2.1 + PKCE + DCR endpoints (for Claude Custom Connectors)
registerOAuthRoutes(app, resolvePublicBaseUrl);

// Health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true });
});

// Download generated PDF by cryptographic token
app.get('/api/exports/:token/:filename?', async (req, res) => {
  const pdfBuffer = await readExportPdf(req.params.token);
  if (!pdfBuffer) {
    return res.status(404).json({ error: 'PDF não encontrado ou expirado.' });
  }
  const safeName = sanitizePdfFilename(req.params.filename || 'etiquetas_pedido');
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
  res.setHeader('Cache-Control', 'no-store');
  return res.send(pdfBuffer);
});

// Catalog search / list (API Bearer or UI)
app.get('/api/catalog', requireApiOrUiAuth, async (req, res) => {
  try {
    const catalog = await getCatalog();
    const query = String(req.query.q || '');
    const categoryId = String(req.query.category || 'all');
    const limit = Number(req.query.limit || 200);
    const labels = searchCatalogLabels(catalog, { query, categoryId, limit });
    res.json({
      totalLabels: catalog.labels.length,
      syncedAt: catalog.syncedAt,
      syncSource: catalog.syncSource,
      categories: catalog.categories,
      labels
    });
  } catch (err) {
    res.status(500).json({ error: err.message || 'Erro ao consultar catálogo.' });
  }
});

// Sync browser catalog to server
app.post('/api/catalog/sync', requireApiOrUiAuth, async (req, res) => {
  try {
    const result = await syncCatalog(req.body, 'browser-sync');
    res.json({ ok: true, ...result });
  } catch (err) {
    res.status(400).json({ error: err.message || 'Payload inválido para sincronização.' });
  }
});

// Generate A4 PDF order via HTTP API
app.post('/api/orders/pdf', requireApiOrUiAuth, async (req, res) => {
  try {
    const { items, filename } = req.body || {};
    const baseUrl = resolvePublicBaseUrl(req);
    const result = await createOrderPdfAndDeliver({
      items,
      filename,
      baseUrl
    });
    res.json({ ok: true, ...result });
  } catch (err) {
    res.status(400).json({ error: err.message || 'Erro ao gerar pedido PDF.' });
  }
});

// Integration status for UI panel (never exposes raw secrets)
app.get('/api/integrations/status', requireTrustedUi, async (req, res) => {
  try {
    const [catalog, auth] = await Promise.all([
      getCatalog(),
      getAuthStatus()
    ]);
    const baseUrl = resolvePublicBaseUrl(req);
    res.json({
      baseUrl,
      mcpUrl: `${baseUrl}/mcp`,
      catalog: {
        labelCount: catalog.labels.length,
        categoryCount: catalog.categories.length,
        syncedAt: catalog.syncedAt,
        syncSource: catalog.syncSource
      },
      auth
    });
  } catch (err) {
    res.status(500).json({ error: err.message || 'Erro ao carregar status de integrações.' });
  }
});

// Claim or rotate MCP token (shown once to UI)
app.post('/api/integrations/mcp-token', requireTrustedUi, async (_req, res) => {
  try {
    const result = await claimOrRotateMcpToken();
    res.json({ ok: true, ...result });
  } catch (err) {
    res.status(400).json({ error: err.message || 'Não foi possível gerar o token MCP.' });
  }
});

// MCP Streamable HTTP endpoint (stateless per-request for remote connectors & mcp-remote)
app.all('/mcp', requireBearerAuth, async (req, res) => {
  const baseUrl = resolvePublicBaseUrl(req);
  const server = createEtiquetaFitMcpServer(baseUrl);
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined
  });

  res.on('close', () => {
    transport.close().catch(() => {});
    server.close().catch(() => {});
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error('Erro no endpoint MCP:', err);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Erro interno ao processar requisição MCP.' });
    }
  }
});

app.use(
  express.static(__dirname, {
    etag: false,
    lastModified: false,
    setHeaders: (res) => {
      res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
    }
  })
);

app.get('*', (_req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`EtiquetaFit server running at http://${HOST}:${PORT}`);
});
