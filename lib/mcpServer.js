import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { getCategoryCounts, LABEL_FORMATS } from './catalog.js';
import { buildOrderPdfBuffer } from './pdfServer.js';
import {
  getCatalog,
  resolveOrderItems,
  saveExportPdf,
  searchCatalogLabels
} from './serverStore.js';

export function resolvePublicBaseUrl(req) {
  const configured = (process.env.PUBLIC_APP_URL || '')
    .trim()
    .replace(/\/mcp\/?$/i, '')
    .replace(/\/+$/, '');
  if (configured) return configured;

  if (!req) return 'http://localhost:3000';
  const protoHeader = (req.get('x-forwarded-proto') || req.protocol || 'http').split(',')[0].trim();
  const hostHeader = (req.get('x-forwarded-host') || req.get('host') || 'localhost:3000').split(',')[0].trim();
  return `${protoHeader}://${hostHeader}`;
}

export async function createOrderPdfAndDeliver({
  items,
  filename,
  baseUrl
}) {
  const catalog = await getCatalog();
  if (!catalog.labels || catalog.labels.length === 0) {
    throw new Error('O catálogo do servidor está vazio. Sincronize o catálogo no app EtiquetaFit primeiro.');
  }

  const resolved = resolveOrderItems(catalog, items);
  if (resolved.items.length === 0) {
    const missingText = resolved.unresolved.join(', ');
    throw new Error(
      missingText
        ? `Nenhuma etiqueta encontrada para: ${missingText}. Use buscar_etiquetas para conferir os nomes disponíveis.`
        : 'Nenhuma etiqueta válida selecionada para o pedido.'
    );
  }

  const pdfBuffer = await buildOrderPdfBuffer(resolved.items, catalog.categories);
  const { token, filename: safeFilename } = await saveExportPdf(pdfBuffer, filename);
  const cleanBase = String(baseUrl || 'http://localhost:3000').replace(/\/+$/, '');
  const downloadUrl = `${cleanBase}/api/exports/${token}/${encodeURIComponent(safeFilename)}`;

  const lineItems = resolved.items.map(item => ({
    id: item.label.id,
    name: item.label.name,
    format: item.format,
    sheets: item.sheets,
    labelsPerSheet: item.labelsPerSheet,
    totalStickers: item.totalStickers
  }));

  return {
    downloadUrl,
    filename: safeFilename,
    summary: resolved.summary,
    items: lineItems,
    unresolved: resolved.unresolved
  };
}

export function createEtiquetaFitMcpServer(baseUrl) {
  const server = new McpServer({
    name: 'EtiquetaFit',
    version: '1.0.0'
  });

  server.registerTool(
    'listar_categorias',
    {
      title: 'Listar Categorias do Catálogo',
      description: 'Lista todas as categorias disponíveis no EtiquetaFit, seus formatos padrão (6x3 cm ou 10x4 cm) e a quantidade de etiquetas em cada categoria.',
      inputSchema: {}
    },
    async () => {
      const catalog = await getCatalog();
      const counts = getCategoryCounts(catalog.labels, catalog.categories);
      const list = catalog.categories.map(c => ({
        id: c.id,
        nome: c.name,
        formatoPadrao: c.defaultFormat || '6x3',
        quantidade: counts[c.id] || 0
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                totalEtiquetas: catalog.labels.length,
                ultimaSincronizacao: catalog.syncedAt,
                categorias: list
              },
              null,
              2
            )
          }
        ]
      };
    }
  );

  server.registerTool(
    'buscar_etiquetas',
    {
      title: 'Buscar Etiquetas',
      description: 'Pesquisa etiquetas no catálogo do EtiquetaFit por nome, ingrediente ou categoria (sem diferenciar acentos). Use antes de gerar o pedido para confirmar os nomes e IDs exatos.',
      inputSchema: {
        consulta: z.string().optional().describe('Texto de busca, como nome da marmita ou suco (ex.: "strogonoff", "desinberry", "escondidinho").'),
        categoria: z.string().optional().describe('ID ou nome da categoria para filtrar (ex.: "sucos", "low-carb").'),
        limite: z.number().int().min(1).max(100).optional().describe('Máximo de resultados (padrão 25).')
      }
    },
    async ({ consulta = '', categoria = 'all', limite = 25 }) => {
      const catalog = await getCatalog();
      const results = searchCatalogLabels(catalog, {
        query: consulta,
        categoryId: categoria,
        limit: limite
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                encontradas: results.length,
                totalNoCatalogo: catalog.labels.length,
                etiquetas: results
              },
              null,
              2
            )
          }
        ]
      };
    }
  );

  server.registerTool(
    'gerar_pedido_pdf',
    {
      title: 'Gerar PDF A4 do Pedido de Etiquetas',
      description: 'Gera um PDF A4 paisagem pronto para impressão combinando marmitas (6x3 cm, 24 por folha) e sucos (10x4 cm, 14 por folha sem espaçamento) e retorna o link direto para download no chat.',
      inputSchema: {
        itens: z
          .array(
            z.object({
              etiquetaId: z.string().optional().describe('ID exato da etiqueta obtido via buscar_etiquetas.'),
              nome: z.string().optional().describe('Nome da etiqueta caso o ID não seja informado.'),
              folhas: z.number().int().min(1).max(50).describe('Quantidade de folhas A4 desejadas para esta etiqueta.')
            })
          )
          .min(1)
          .describe('Lista de etiquetas e quantidades de folhas A4.'),
        nomeArquivo: z.string().optional().describe('Nome opcional do arquivo PDF gerado.')
      }
    },
    async ({ itens, nomeArquivo }) => {
      const result = await createOrderPdfAndDeliver({
        items: itens,
        filename: nomeArquivo,
        baseUrl
      });

      const lines = [
        `PDF gerado com sucesso: ${result.filename}`,
        `Link direto para baixar: ${result.downloadUrl}`,
        `Resumo: ${result.summary.totalModels} modelo(s), ${result.summary.totalSheets} folha(s) A4, ${result.summary.totalStickers} etiqueta(s).`,
        '',
        'Itens incluídos:',
        ...result.items.map(
          i => `- ${i.sheets} folha(s) de ${i.name} (${LABEL_FORMATS[i.format].shortDim} · ${i.totalStickers} etiquetas)`
        )
      ];

      if (result.unresolved.length > 0) {
        lines.push('', `Não encontrados no catálogo: ${result.unresolved.join(', ')}`);
      }

      return {
        content: [
          {
            type: 'text',
            text: lines.join('\n')
          }
        ]
      };
    }
  );

  return server;
}
